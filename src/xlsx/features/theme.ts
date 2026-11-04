// Default Office theme part (xl/theme/theme1.xml). A feature so that `./lite` bundles can leave
// it out (~0.5 KB): files without a theme part open fine in Excel / LibreOffice (theme colors
// and the Calibri font scheme fall back to the application's Office defaults). The package root
// always enables it; `WriteOptions.theme = false` removes it. Reading needs no half: a theme in
// the source file is kept raw by preservation and replaces this one.

import type { FeatureWriter } from './types.js';

const CT = 'application/vnd.openxmlformats-officedocument.theme+xml';

/** Minimal valid Office theme: the default palette and Calibri fonts, placeholder-only format scheme. */
function themeXml(): string {
  const fill = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
  const x3 = (s: string): string => s + s + s;
  const clr = (n: string, v: string): string => `<a:${n}><a:srgbClr val="${v}"/></a:${n}>`;
  const font = (n: string, f: string): string =>
    `<a:${n}Font><a:latin typeface="${f}"/><a:ea typeface=""/><a:cs typeface=""/></a:${n}Font>`;
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office">' +
    '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>' +
    clr('dk2', '44546A') +
    clr('lt2', 'E7E6E6') +
    clr('accent1', '4472C4') +
    clr('accent2', 'ED7D31') +
    clr('accent3', 'A5A5A5') +
    clr('accent4', 'FFC000') +
    clr('accent5', '5B9BD5') +
    clr('accent6', '70AD47') +
    clr('hlink', '0563C1') +
    clr('folHlink', '954F72') +
    `</a:clrScheme><a:fontScheme name="Office">${font('major', 'Calibri Light')}${font('minor', 'Calibri')}</a:fontScheme>` +
    `<a:fmtScheme name="Office"><a:fillStyleLst>${x3(fill)}</a:fillStyleLst><a:lnStyleLst>${x3(`<a:ln w="9525">${fill}</a:ln>`)}</a:lnStyleLst>` +
    `<a:effectStyleLst>${x3('<a:effectStyle><a:effectLst/></a:effectStyle>')}</a:effectStyleLst><a:bgFillStyleLst>${x3(fill)}</a:bgFillStyleLst></a:fmtScheme>` +
    '</a:themeElements></a:theme>'
  );
}

/**
 * The default theme part. Writes nothing when the workbook kept a theme from the file it was
 * read from (never two themes).
 */
export const themeWriter: FeatureWriter = {
  name: 'theme',
  book: {
    write(ctx) {
      const pb = ctx.workbook.preserved;
      if (pb?.implicit.some((k) => /\/theme$/.test(pb.rels[k]!.type))) return;
      ctx.addPart('xl/theme/theme1.xml', themeXml(), CT);
      ctx.addRel('theme', 'theme/theme1.xml');
    },
  },
};
