---
"turbosheet": minor
---

Threaded comments and custom views (group 18, #111, #112):

- **Threaded comments.** Excel 365 conversations are now modelled instead of kept raw: `ws.threadedComments` (ref, text, author, date, done, replies) and `wb.persons` (authors, with their account ids kept). `addThreadedComment(ws, ref, { text, author, replies })` adds one. The writer emits the threaded part, the persons part and the legacy note Excel expects for each thread. The reader turns those generated notes back into the thread, so `ws.comments` holds only real notes. Files round-trip byte for byte.
- **Custom views.** `wb.customViews` and `ws.customSheetViews` (View > Custom Views) are read and written. `addCustomView(wb, name)` captures every sheet as it is now. New feature halves: `customViewsWriter` / `customViewsReader`.
- Framework: `BookEmitter.owns(wb)` lets a feature own a workbook element, so the copy kept from the read file is not written back.

Checked in Excel for Mac: new threads and views open without a repair prompt, and a file Excel saved with them round-trips.
