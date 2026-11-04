// Shared string table (sst). Writers intern every string cell to get its index; readers
// push strings in file order and look them up with get().

export class SharedStrings {
  private readonly _list: string[] = [];
  private readonly _index = new Map<string, number>();
  private _count = 0;

  /** String → index (deduped). Every call counts toward `count`. */
  intern(s: string): number {
    this._count++;
    let i = this._index.get(s);
    if (i === undefined) {
      i = this._list.length;
      this._list.push(s);
      this._index.set(s, i);
    }
    return i;
  }

  /** Reader hook: append a string in file order without dedupe. Returns its index. */
  push(s: string): number {
    const i = this._list.length;
    this._list.push(s);
    if (!this._index.has(s)) this._index.set(s, i);
    return i;
  }

  get(i: number): string | undefined {
    return this._list[i];
  }

  /** Total references (the sst `count` attribute). */
  get count(): number {
    return this._count;
  }

  /** Distinct strings (the sst `uniqueCount` attribute). */
  get uniqueCount(): number {
    return this._list.length;
  }

  /** All strings in index order. */
  get strings(): readonly string[] {
    return this._list;
  }
}
