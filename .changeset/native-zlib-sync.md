---
"turbosheet": minor
---

`writeXlsx` and `readXlsx` use native zlib in Node, Bun and Deno (looked up at runtime with `process.getBuiltinModule`; browsers keep fflate). `writeXlsx` deflates with `zlib.deflateRawSync` in ~1 MiB sync-flushed batches: 200k x 35 balanced 10.9 s to 7.3 s with ~1.5% smaller files; `max-speed` 3.8 s to 1.7 s. `readXlsx` inflates parts up to 64 MiB with `zlib.inflateRawSync` (5.8 s to 4.8 s on the same file), with zip-bomb limits enforced through `maxOutputLength`; larger parts still stream through fflate. Closes #75.
