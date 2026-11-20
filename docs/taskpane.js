// Plagx Draft Preflight task pane for Word: reads the document or selection, runs Plagx
// Preflight in the task pane and lists the issues. Clicking an issue selects it in the
// document. Nothing is stored or sent anywhere.
import { checkDraft } from './lib/index.js';
import { readDocument, selectText } from './word.js';

const LEVELS = { error: 'Error', warning: 'Warning', info: 'Suggestion' };
const $ = (id) => document.getElementById(id);
let lastText = '';

const count = (n, word) => `${n} ${word}${n === 1 ? '' : 's'}`;
const plain = (msg) => {
  const text = msg.replace(/^Suggestion: /, '');
  return text.charAt(0).toUpperCase() + text.slice(1);
};

function el(tag, className, text) {
  const node = document.createElement(tag);
  if (className) node.className = className;
  if (text !== undefined) node.textContent = text;
  return node;
}

const message = (text, className = 'plagx-hint') => $('result').replaceChildren(el('p', className, text));

/** How many times `snippet` appears in the checked text before `start`. */
function occurrenceBefore(snippet, start) {
  let n = 0;
  for (let i = lastText.indexOf(snippet); i !== -1 && i < start; i = lastText.indexOf(snippet, i + 1)) n++;
  return n;
}

async function locate(issue, item) {
  const snippet = issue.text.split('\n')[0];
  item.classList.add('is-busy');
  try {
    if (!(await selectText(window.Word, snippet, occurrenceBefore(snippet, issue.start)))) item.classList.add('is-missing');
  } catch {
    item.classList.add('is-missing');
  } finally {
    item.classList.remove('is-busy');
  }
}

function render(result) {
  const out = $('result');
  out.replaceChildren();
  const { stats } = result;
  const tiles = el('div', 'plagx-stats');
  for (const [value, label] of [[stats.words, 'Words'], [stats.readingTimeMinutes, 'Min read'], [stats.sentences, 'Sentences'], [stats.paragraphs, 'Paragraphs']]) {
    const tile = el('div', 'plagx-stat');
    tile.append(el('strong', '', String(value)), el('span', '', label));
    tiles.append(tile);
  }
  const pills = el('div', 'plagx-pills');
  pills.append(
    el('span', `plagx-pill${result.errors ? ' is-error' : ''}`, count(result.errors, 'error')),
    el('span', `plagx-pill${result.warnings ? ' is-warning' : ''}`, count(result.warnings, 'warning')),
    el('span', `plagx-pill${result.infos ? ' is-info' : ''}`, count(result.infos, 'suggestion')),
  );
  out.append(tiles, pills);
  if (!result.issues.length) {
    out.append(el('p', 'plagx-done', '✓ No issues found.'));
    return;
  }
  const list = el('ul', 'plagx-issues');
  for (const issue of result.issues) {
    const item = el('li', `plagx-issue is-${issue.level}`);
    item.append(el('span', 'plagx-issue__level', LEVELS[issue.level]), el('span', 'plagx-issue__message', plain(issue.message)));
    if (issue.text) {
      item.append(el('span', 'plagx-issue__text', issue.text.length > 80 ? `${issue.text.slice(0, 80)}…` : issue.text));
      item.tabIndex = 0;
      item.setAttribute('role', 'button');
      item.title = 'Find in document';
      item.addEventListener('click', () => locate(issue, item));
      item.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' || e.key === ' ') {
          e.preventDefault();
          locate(issue, item);
        }
      });
    }
    list.append(item);
  }
  out.append(list);
}

function options() {
  const number = (id) => Number.parseInt($(id).value, 10) || undefined;
  return {
    minWords: number('min-words'),
    maxWords: number('max-words'),
    style: $('style').value,
    checks: [...document.querySelectorAll('input[name="check"]')].filter((b) => b.checked).map((b) => b.value),
  };
}

async function run() {
  const fromSelection = document.querySelector('input[name="source"]:checked').value === 'selection';
  const button = $('check');
  button.disabled = true;
  message('Reading the document…');
  try {
    await window.Office?.onReady();
    const text = await readDocument(window.Word, fromSelection);
    if (!text) {
      message(fromSelection ? 'Select some text in the document, then check again.' : 'This document has no text to check yet.');
      return;
    }
    lastText = text;
    render(checkDraft(text, options()));
  } catch (err) {
    message(err.message, 'plagx-fail');
  } finally {
    button.disabled = false;
  }
}

$('check').addEventListener('click', run);
if (!window.Office) message('Open this add-in in Microsoft Word to check a document.');
