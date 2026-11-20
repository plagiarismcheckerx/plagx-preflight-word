import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { joinParagraphs, paragraphText, readDocument, selectText } from '../docs/word.js';
import { checkDraft } from '../docs/lib/index.js';

/** A stand-in for Word.run over a list of paragraphs and an optional selection. */
function fakeWord(paragraphs, { selection = [], selectedText = '' } = {}) {
  const selected = [];
  const collection = (items) => ({ items, load() {} });
  const Word = {
    selected,
    run: async (fn) => fn({
      sync: async () => {},
      document: {
        body: {
          paragraphs: collection(paragraphs),
          search(needle) {
            const hits = [];
            paragraphs.forEach((p, i) => {
              for (let at = p.text.indexOf(needle); at !== -1; at = p.text.indexOf(needle, at + 1)) {
                hits.push({ select: () => selected.push([i, at]) });
              }
            });
            return collection(hits);
          },
        },
        getSelection: () => ({ paragraphs: collection(selection), text: selectedText, load() {} }),
      },
    }),
  };
  return Word;
}

test('paragraph text: headings marked, control characters removed, empty skipped', () => {
  assert.equal(paragraphText({ text: 'References\r', styleBuiltIn: 'Heading1' }), '# References');
  assert.equal(paragraphText({ text: 'My essay', styleBuiltIn: 'Title' }), '# My essay');
  assert.equal(paragraphText({ text: 'Body\u000btext', styleBuiltIn: 'Normal' }), 'Body text');
  assert.equal(paragraphText({ text: '   ', styleBuiltIn: 'Normal' }), '');
  assert.equal(joinParagraphs([{ text: 'A' }, { text: '' }, { text: 'B', styleBuiltIn: 'Heading2' }]), 'A\n\n# B');
});

const DOC = [
  { text: 'Intro', styleBuiltIn: 'Heading1' },
  { text: 'This is is a test (Smith, 2020).', styleBuiltIn: 'Normal' },
  { text: 'It is is fine.', styleBuiltIn: 'Normal' },
  { text: 'References', styleBuiltIn: 'Heading1' },
  { text: 'Lee, K. (2019). Planning.', styleBuiltIn: 'Normal' },
];

test('reads the whole document, a multi-paragraph selection, or selected text in one paragraph', async () => {
  const text = await readDocument(fakeWord(DOC));
  assert.match(text, /^# Intro\n\nThis is is/);
  const codes = checkDraft(text).issues.map((i) => i.code);
  assert.ok(codes.includes('repeated-word') && codes.includes('citation-no-reference') && codes.includes('reference-not-cited'));
  assert.equal(await readDocument(fakeWord(DOC, { selection: DOC.slice(0, 2) }), true), '# Intro\n\nThis is is a test (Smith, 2020).');
  assert.equal(await readDocument(fakeWord(DOC, { selection: [DOC[1]], selectedText: 'is is a test' }), true), 'is is a test');
});

test('selects the right occurrence, falls back to the first, and reports misses', async () => {
  const Word = fakeWord(DOC);
  assert.equal(await selectText(Word, 'is is', 1), true);
  assert.deepEqual(Word.selected.at(-1), [1, 5]);
  assert.equal(await selectText(Word, 'is is', 9), true);
  assert.deepEqual(Word.selected.at(-1), [1, 2]);
  assert.equal(await selectText(Word, 'nowhere', 0), false);
  assert.equal(await selectText(Word, '', 0), false);
});

test('manifest points every URL at the GitHub Pages site and the files exist', () => {
  const manifest = readFileSync(new URL('../manifest.xml', import.meta.url), 'utf8');
  const base = 'https://plagiarismcheckerx.github.io/plagx-preflight-word/';
  for (const [, url] of manifest.matchAll(/DefaultValue="(https:\/\/plagiarismcheckerx\.github\.io[^"]+)"/g)) {
    assert.ok(url.startsWith(base), url);
    readFileSync(new URL(`../docs/${url.slice(base.length)}`, import.meta.url));
  }
});
