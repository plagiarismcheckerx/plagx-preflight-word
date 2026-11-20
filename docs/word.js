// Reads the Word document (or selection) as plain text and selects text in it, with Office.js.
// Headings become "# Heading" so Plagx Preflight can find the References section.

/** "# Text" for Title and Heading styles, the trimmed text otherwise; '' for empty paragraphs. */
export function paragraphText(p) {
  const text = (p.text || '').replace(/[\r\u000b\u0007]/g, ' ').trim();
  if (!text) return '';
  const style = String(p.styleBuiltIn || '');
  return /^(Heading\d|Title)$/.test(style) ? `# ${text}` : text;
}

export const joinParagraphs = (items) => items.map(paragraphText).filter(Boolean).join('\n\n');

/** Text of the whole body, or of the selection when `selection` is true. */
export async function readDocument(Word, selection = false) {
  return Word.run(async (context) => {
    const range = selection ? context.document.getSelection() : context.document.body;
    const paragraphs = range.paragraphs;
    paragraphs.load('items/text,items/styleBuiltIn');
    await context.sync();
    if (selection && paragraphs.items.length <= 1) {
      // A selection inside one paragraph: use just the selected text.
      const sel = context.document.getSelection();
      sel.load('text');
      await context.sync();
      return paragraphText({ text: sel.text });
    }
    return joinParagraphs(paragraphs.items);
  });
}

/** Selects the `occurrence`-th (0-based) match of `snippet`; returns false if not found. */
export async function selectText(Word, snippet, occurrence = 0) {
  // Word search is limited to 255 characters, and "^" starts a special code such as ^p.
  const needle = String(snippet || '').split('\n')[0].trim().slice(0, 200).replace(/\^/g, '^^');
  if (!needle) return false;
  return Word.run(async (context) => {
    const results = context.document.body.search(needle, { matchCase: true });
    results.load('items');
    await context.sync();
    if (!results.items.length) return false;
    const hit = results.items[occurrence] || results.items[0];
    hit.select();
    await context.sync();
    return true;
  });
}
