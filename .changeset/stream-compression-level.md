---
"turbosheet": minor
---

Streaming writers (`createXlsxWriter`, `xlsxStream`) now honor `compression` / `profile` in Node, Bun and Deno: entries are deflated off the main thread by `node:zlib` at the requested level (looked up at runtime with `process.getBuiltinModule`, no static `node:*` import), instead of `CompressionStream`, which has no level. `max-speed` streaming is ~2.5x faster (300k rows: 1.98 s to 0.76 s). Zip CRC-32 uses `zlib.crc32` when available (~5x faster than the JS table), so the default level is ~15% faster too. Browsers keep `CompressionStream` (platform default level, now documented). Closes #71.
