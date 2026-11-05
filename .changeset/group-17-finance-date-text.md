---
"turbosheet": minor
---

Formula functions, group 17 (#109, #110): 50 new functions, 213 in total. The results were checked against Excel for Mac.

- **Financial:** PMT, IPMT, PPMT, PV, FV, NPER, NPV, IRR, XNPV, XIRR, RATE, SLN, DB, DDB, EFFECT, NOMINAL. IRR / XIRR / RATE use Excel's Newton iterations and defaults.
- **Date / time:** TIME, HOUR, MINUTE, SECOND, DATEVALUE, TIMEVALUE, WEEKNUM, ISOWEEKNUM, NETWORKDAYS(.INTL), WORKDAY(.INTL), YEARFRAC.
- **Text:** CHAR, CODE, UNICHAR, UNICODE, CLEAN, FIXED, DOLLAR, NUMBERVALUE, T, TEXTBEFORE, TEXTAFTER, TEXTSPLIT (spills).
- **Information:** N, TYPE, ISEVEN, ISODD, ISREF, ISFORMULA, ERROR.TYPE, ADDRESS, HYPERLINK (returns the value only).
