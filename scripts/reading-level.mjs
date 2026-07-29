// Flesch-Kincaid grade level for a markdown file.
//
// Exists because docs/setup-guide.md carries a hard rule - no more than a 6th
// grade reading level - and a reading level that is asserted rather than
// measured is reliably three grades optimistic. Someone fluent in a subject
// cannot feel the difficulty of their own prose.
//
// Exits 1 above the threshold, so it is a gate rather than a report.
//
//   node scripts/reading-level.mjs docs/setup-guide.md
//   node scripts/reading-level.mjs docs/setup-guide.md 8      (custom max)

import { readFileSync } from 'node:fs'

// Heuristic syllable count. Wrong on some words; unbiased enough across a
// document for the aggregate to be meaningful, which is all the formula needs.
const syllables = (w) => {
  w = w.toLowerCase().replace(/[^a-z]/g, '')
  if (w.length <= 3) return w.length > 0 ? 1 : 0
  w = w.replace(/(?:[^laeiouy]es|ed|[^laeiouy]e)$/, '').replace(/^y/, '')
  return (w.match(/[aeiouy]{1,2}/g) || ['x']).length
}

const file = process.argv[2]
const maxGrade = process.argv[3] ? Number(process.argv[3]) : 6.0

if (!file) {
  console.error('usage: node scripts/reading-level.mjs <file.md> [maxGrade]')
  process.exit(2)
}

// Strip everything that is not prose BEFORE measuring. A keystroke like
// "Alt+F11", a file name, or a fenced code block is not a sentence, and
// scoring it drags the result around for reasons that have nothing to do with
// how hard the writing is to read.
const prose = readFileSync(file, 'utf8')
  .replace(/```[\s\S]*?```/g, ' ')      // fenced code
  .replace(/`[^`]*`/g, ' ')             // inline code
  .replace(/\[([^\]]*)\]\([^)]*\)/g, '$1') // links: keep the text, drop the target
  .replace(/<[^>]+>/g, ' ')             // html
  .replace(/^\s{0,3}[#>]+\s*/gm, ' ')   // headings and quotes
  .replace(/^\s*[*+-]\s+/gm, ' ')       // list bullets
  .replace(/^\s*\d+\.\s+/gm, ' ')       // ordered list markers
  .replace(/^\s*\|.*\|\s*$/gm, ' ')     // table rows
  .replace(/\*\*?|__?/g, '')            // emphasis marks

const sentences = (prose.match(/[.!?]+(?=\s|$)/g) || []).length || 1
const words = prose.split(/\s+/).filter((w) => /[a-z]/i.test(w))

if (words.length === 0) {
  console.error(`${file}: no prose found to measure`)
  process.exit(2)
}

const syl = words.reduce((n, w) => n + syllables(w), 0)
const grade = 0.39 * (words.length / sentences) + 11.8 * (syl / words.length) - 15.59

const verdict = grade > maxGrade ? 'OVER' : 'ok'
console.log(
  `${file}: grade ${grade.toFixed(2)} (max ${maxGrade.toFixed(1)}) - ${verdict}\n` +
    `  ${words.length} words, ${sentences} sentences, ` +
    `${(words.length / sentences).toFixed(1)} words/sentence, ` +
    `${(syl / words.length).toFixed(2)} syllables/word`
)

if (grade > maxGrade) {
  console.error(
    '\nToo hard. Shorten sentences and prefer shorter words.\n' +
      'Do NOT delete steps to pass - a guide that passes by being incomplete has failed.'
  )
  process.exit(1)
}
