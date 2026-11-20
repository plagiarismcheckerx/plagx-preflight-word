# Plagx Draft Preflight for Word

*A free academic writing tool by Plagiarism Checker X.*

**Plagx Draft Preflight is a free, open-source Microsoft Word add-in that checks your draft before you submit it.** Open it from **Home → Check draft** and check the whole document or a selection for word count and reading time, in-text citations against the reference list, punctuation and basic grammar. Click an issue to jump to it in the document. Use it as an **essay checker**, **word counter**, **citation checker** or **reference list checker** for essays, papers and articles.

- Checks run in the task pane, in your browser: nothing is stored or sent anywhere
- Reads the document only when you click Check draft; never changes your text
- APA and MLA in-text citations
- Works in Word on Windows, Mac, the web and iPad
- Same rules as the [plagx-preflight](https://github.com/plagiarismcheckerx/plagx-preflight) library

**[Try the live demo →](https://plagiarismcheckerx.github.io/plagx-preflight/)**

## Use

1. Install **Plagx Draft Preflight** from Microsoft Marketplace (in Word: **Home → Add-ins**).
2. Click **Home → Check draft**.
3. Choose **Whole document** or **Selected text**, set word limits and citation style if you need them, and click **Check draft**.

Use a heading style (Heading 1–6 or Title) for your References or Works Cited heading, with each reference in its own paragraph.

## Checks

| Group | Examples |
|---|---|
| Length & structure | Word, sentence and paragraph counts, reading time, word limits, long sentences and paragraphs |
| References | Citations with no entry in the reference list, references that are never cited |
| Punctuation | Unmatched quotes and brackets, double spaces, spacing around punctuation |
| Basic grammar | Repeated words, *a* / *an*, sentence capitals, common word confusions; optional passive-voice hints |

## Privacy

The text is checked in the task pane and discarded when you close it. No accounts, tracking or network requests with your content. See the [privacy policy](https://plagiarismcheckerx.github.io/plagx-preflight-word/privacy.html) and [terms of use](https://plagiarismcheckerx.github.io/plagx-preflight-word/terms.html).

## Development

The add-in is a static site in `docs/`, served by GitHub Pages at <https://plagiarismcheckerx.github.io/plagx-preflight-word/>. There is no build step.

```bash
npm install
npm test           # document reading, text selection and manifest tests
npm run validate   # check manifest.xml with Microsoft's validator
npm run sync       # copy the latest core from ../plagx-preflight/src
```

To try it in Word, [sideload](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins-for-testing) `manifest.xml` (Word on the web: **Home → Add-ins → More add-ins → My add-ins → Upload my add-in**).

| Path | What it does |
|---|---|
| `manifest.xml` | Office add-in manifest |
| `docs/taskpane.html`, `docs/taskpane.js`, `docs/styles.css` | The task pane |
| `docs/word.js` | Reads the document or selection and selects an issue's text (Office.js) |
| `docs/lib/` | The Plagx Preflight core, copied from [plagx-preflight](https://github.com/plagiarismcheckerx/plagx-preflight) (MIT) |
| `docs/index.html`, `privacy.html`, `terms.html` | Support, privacy and terms pages |

## Plagx Preflight - Free Draft & Grammar Checker

An open-source writing utility for students, researchers and developers. Check length, structure, citations, punctuation and basic grammar before you submit.

Developed by Plagiarism Checker X.

Plagx Preflight gives writing suggestions; it does not detect plagiarism or AI-generated text. For full document checks, explore [Text Similarity Detector](https://plagiarismcheckerx.com/?utm_source=plagx-preflight-word-readme) and its [AI Content Detector](https://plagiarismcheckerx.com/ai-content-detector?utm_source=plagx-preflight-word-readme).

## License

MIT © Plagiarism Checker X
