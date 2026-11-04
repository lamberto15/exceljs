---
"turbosheet": patch
---

Streaming: documented and tested backpressure (`addRows` / `flush` wait for slow consumers); regression tests for reported ExcelJS streaming bugs (pictures and number formats in the streaming writer, multi-byte text split across read chunks, constant memory, merges over tables, background picture with notes).
