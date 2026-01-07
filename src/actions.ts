import type {
  App,
  Editor,
  EditorChange,
  EditorPosition,
  EditorSelection,
} from 'obsidian';
import {
  CASE,
  SEARCH_DIRECTION,
  MATCHING_BRACKETS,
  MATCHING_QUOTES,
  MATCHING_QUOTES_BRACKETS,
  MatchingCharacterMap,
  CODE_EDITOR,
  LIST_CHARACTER_REGEX,
} from './constants';
import { SettingsState } from './state';
import {
  CheckCharacter,
  EditorActionCallbackNewArgs,
  findAllMatchPositions,
  findNextMatchPosition,
  findPosOfNextCharacter,
  formatRemainingListPrefixes,
  getLeadingWhitespace,
  getLineEndPos,
  getLineStartPos,
  getNextCase,
  toTitleCase,
  getSelectionBoundaries,
  wordRangeAtPos,
  getSearchText,
  getNextListPrefix,
  isNumeric,
} from './utils';

export const insertLineAbove = (
  editor: Editor,
  selection: EditorSelection,
  args: EditorActionCallbackNewArgs,
) => {
  const { line } = selection.head;
  const startOfCurrentLine = getLineStartPos(line);

  const contentsOfCurrentLine = editor.getLine(line);
  const indentation = getLeadingWhitespace(contentsOfCurrentLine);

  let listPrefix = '';
  if (
    SettingsState.autoInsertListPrefix &&
    line > 0 &&
    // If inside a list, only insert prefix if within the same list
    editor.getLine(line - 1).trim().length > 0
  ) {
    listPrefix = getNextListPrefix(contentsOfCurrentLine, 'before');
    if (isNumeric(listPrefix)) {
      formatRemainingListPrefixes(editor, line, indentation);
    }
  }

  const changes: EditorChange[] = [
    { from: startOfCurrentLine, text: indentation + listPrefix + '\n' },
  ];
  const newSelection = {
    from: {
      ...startOfCurrentLine,
      // Offset by iteration
      line: startOfCurrentLine.line + args.iteration,
      ch: indentation.length + listPrefix.length,
    },
  };
  return {
    changes,
    newSelection,
  };
};

export const insertLineBelow = (
  editor: Editor,
  selection: EditorSelection,
  args: EditorActionCallbackNewArgs,
) => {
  const { line } = selection.head;
  const startOfCurrentLine = getLineStartPos(line);
  const endOfCurrentLine = getLineEndPos(line, editor);

  const contentsOfCurrentLine = editor.getLine(line);
  const indentation = getLeadingWhitespace(contentsOfCurrentLine);

  let listPrefix = '';
  if (SettingsState.autoInsertListPrefix) {
    listPrefix = getNextListPrefix(contentsOfCurrentLine, 'after');

    // Performing this action on an empty list item should delete it
    if (listPrefix === null) {
      const changes: EditorChange[] = [
        { from: startOfCurrentLine, to: endOfCurrentLine, text: '' },
      ];
      const newSelection = {
        from: {
          line,
          ch: 0,
        },
      };
      return {
        changes,
        newSelection,
      };
    }

    if (isNumeric(listPrefix)) {
      formatRemainingListPrefixes(editor, line + 1, indentation);
    }
  }

  const changes: EditorChange[] = [
    { from: endOfCurrentLine, text: '\n' + indentation + listPrefix },
  ];
  const newSelection = {
    from: {
      // Offset by iteration
      line: line + 1 + args.iteration,
      ch: indentation.length + listPrefix.length,
    },
  };
  return {
    changes,
    newSelection,
  };
};

// Note: don't use the built-in exec method for 'deleteLine' as there is a bug
// where running it on a line that is long enough to be wrapped will focus on
// the previous line instead of the next line after deletion
let numLinesDeleted = 0;
export const deleteLine = (
  editor: Editor,
  selection: EditorSelection,
  args: EditorActionCallbackNewArgs,
) => {
  const { from, to, hasTrailingNewline } = getSelectionBoundaries(selection);

  if (to.line === editor.lastLine()) {
    // There is no 'next line' when cursor is on the last line
    const previousLine = Math.max(0, from.line - 1);
    const endOfPreviousLine = getLineEndPos(previousLine, editor);
    const changes: EditorChange[] = [
      {
        from: from.line === 0 ? getLineStartPos(0) : endOfPreviousLine,
        to:
          // Exclude line starting at trailing newline at end of document from being deleted
          to.ch === 0
            ? getLineStartPos(to.line)
            : getLineEndPos(to.line, editor),
        text: '',
      },
    ];
    const newSelection = {
      from: {
        line: previousLine,
        ch: Math.min(from.ch, endOfPreviousLine.ch),
      },
    };
    return {
      changes,
      newSelection,
    };
  }

  // Reset offset at the start of a new bulk delete operation
  if (args.iteration === 0) {
    numLinesDeleted = 0;
  }
  // Exclude line starting at trailing newline from being deleted
  const toLine = hasTrailingNewline ? to.line - 1 : to.line;
  const endOfNextLine = getLineEndPos(toLine + 1, editor);
  const changes: EditorChange[] = [
    {
      from: getLineStartPos(from.line),
      to: getLineStartPos(toLine + 1),
      text: '',
    },
  ];
  const newSelection = {
    from: {
      // Offset by the number of lines deleted in all previous iterations
      line: from.line - numLinesDeleted,
      ch: Math.min(to.ch, endOfNextLine.ch),
    },
  };
  // This needs to be calculated after setting the new selection as it only
  // applies for subsequent iterations
  numLinesDeleted += toLine - from.line + 1;
  return {
    changes,
    newSelection,
  };
};

export const deleteToStartOfLine = (
  editor: Editor,
  selection: EditorSelection,
) => {
  const pos = selection.head;
  let startPos = getLineStartPos(pos.line);

  if (pos.line === 0 && pos.ch === 0) {
    // We're at the start of the document so do nothing
    return selection;
  }

  if (pos.line === startPos.line && pos.ch === startPos.ch) {
    // We're at the start of the line so delete the preceding newline
    startPos = getLineEndPos(pos.line - 1, editor);
  }

  editor.replaceRange('', startPos, pos);
  return {
    anchor: startPos,
  };
};

export const deleteToEndOfLine = (
  editor: Editor,
  selection: EditorSelection,
) => {
  const pos = selection.head;
  let endPos = getLineEndPos(pos.line, editor);

  if (pos.line === endPos.line && pos.ch === endPos.ch) {
    // We're at the end of the line so delete just the newline
    endPos = getLineStartPos(pos.line + 1);
  }

  editor.replaceRange('', pos, endPos);
  return {
    anchor: pos,
  };
};

export const deleteToEndOfSentence = (
  editor: Editor,
  selection: EditorSelection,
) => {
  const pos = selection.head;
  const lineEndPos = getLineEndPos(pos.line, editor);
  const restOfLine = editor.getRange(pos, lineEndPos);

  // Find the next sentence-ending punctuation (.!?) followed by optional closing delimiters and space or end of line
  const sentenceEndMatch = restOfLine.match(/[.!?]["')\]}*_`]*(?=\s|$)/);

  let endPos: EditorPosition;
  if (sentenceEndMatch && sentenceEndMatch.index !== undefined) {
    // Found a sentence end on this line - delete up to and including the punctuation and closing delimiters
    endPos = {
      line: pos.line,
      ch: pos.ch + sentenceEndMatch.index + sentenceEndMatch[0].length,
    };
  } else {
    // No sentence end found on this line - delete to end of line
    endPos = lineEndPos;

    // If we're at the end of the line, delete the newline too
    if (pos.line === endPos.line && pos.ch === endPos.ch) {
      endPos = getLineStartPos(pos.line + 1);
    }
  }

  editor.replaceRange('', pos, endPos);
  return {
    anchor: pos,
  };
};

export const deleteToStartOfSentence = (
  editor: Editor,
  selection: EditorSelection,
) => {
  const pos = selection.head;
  const lineStartPos = getLineStartPos(pos.line);
  const startOfLine = editor.getRange(lineStartPos, pos);

  // Find the last sentence-ending punctuation (.!?) followed by optional closing delimiters and space
  // We search for the punctuation, any closing delimiters, and any following whitespace
  const sentenceStartMatches = Array.from(startOfLine.matchAll(/[.!?]["')\]}*_`]*\s+/g));

  let startPos: EditorPosition;
  if (sentenceStartMatches.length > 0) {
    // Found a sentence end - delete from after the punctuation, closing delimiters, and space to cursor
    const lastMatch = sentenceStartMatches[sentenceStartMatches.length - 1];
    const matchEnd = (lastMatch.index ?? 0) + lastMatch[0].length;
    startPos = {
      line: pos.line,
      ch: lineStartPos.ch + matchEnd,
    };
  } else {
    // No sentence end found - delete to start of line
    startPos = lineStartPos;

    // If we're already at the start of the line, delete the preceding newline
    if (pos.line === startPos.line && pos.ch === startPos.ch) {
      if (pos.line > 0) {
        startPos = getLineEndPos(pos.line - 1, editor);
      } else {
        // We're at the start of the document, do nothing
        return selection;
      }
    }
  }

  editor.replaceRange('', startPos, pos);
  return {
    anchor: startPos,
  };
};

export const joinLines = (editor: Editor, selection: EditorSelection) => {
  const { from, to } = getSelectionBoundaries(selection);
  const { line } = from;

  let endOfCurrentLine = getLineEndPos(line, editor);
  const joinRangeLimit = Math.max(to.line - line, 1);
  const selectionLength = editor.posToOffset(to) - editor.posToOffset(from);
  let trimmedChars = '';

  for (let i = 0; i < joinRangeLimit; i++) {
    if (line === editor.lineCount() - 1) {
      break;
    }
    endOfCurrentLine = getLineEndPos(line, editor);
    const endOfNextLine = getLineEndPos(line + 1, editor);
    const contentsOfCurrentLine = editor.getLine(line);
    const contentsOfNextLine = editor.getLine(line + 1);

    const charsToTrim = contentsOfNextLine.match(LIST_CHARACTER_REGEX) ?? [];
    trimmedChars += charsToTrim[0] ?? '';

    const newContentsOfNextLine = contentsOfNextLine.replace(
      LIST_CHARACTER_REGEX,
      '',
    );
    if (
      newContentsOfNextLine.length > 0 &&
      contentsOfCurrentLine.charAt(endOfCurrentLine.ch - 1) !== ' '
    ) {
      editor.replaceRange(
        ' ' + newContentsOfNextLine,
        endOfCurrentLine,
        endOfNextLine,
      );
    } else {
      editor.replaceRange(
        newContentsOfNextLine,
        endOfCurrentLine,
        endOfNextLine,
      );
    }
  }

  if (selectionLength === 0) {
    return {
      anchor: endOfCurrentLine,
    };
  }
  return {
    anchor: from,
    head: {
      line: from.line,
      ch: from.ch + selectionLength - trimmedChars.length,
    },
  };
};

export const copyLine = (
  editor: Editor,
  selection: EditorSelection,
  direction: 'up' | 'down',
) => {
  const { from, to, hasTrailingNewline } = getSelectionBoundaries(selection);
  const fromLineStart = getLineStartPos(from.line);
  // Exclude line starting at trailing newline from being duplicated
  const toLine = hasTrailingNewline ? to.line - 1 : to.line;
  const toLineEnd = getLineEndPos(toLine, editor);
  const contentsOfSelectedLines = editor.getRange(fromLineStart, toLineEnd);
  if (direction === 'up') {
    editor.replaceRange('\n' + contentsOfSelectedLines, toLineEnd);
    return selection;
  } else {
    editor.replaceRange(contentsOfSelectedLines + '\n', fromLineStart);
    // This uses `to.line` instead of `toLine` to avoid a double adjustment
    const linesSelected = to.line - from.line + 1;
    return {
      anchor: { line: toLine + 1, ch: from.ch },
      head: { line: toLine + linesSelected, ch: to.ch },
    };
  }
};

/*
Properties used to distinguish between selections that are programmatic
(expanding from a cursor selection) vs. manual (using a mouse / Shift + arrow
keys). This controls the match behaviour for selectWordOrNextOccurrence.
*/
let isManualSelection = true;
export const setIsManualSelection = (value: boolean) => {
  isManualSelection = value;
};
export let isProgrammaticSelectionChange = false;
export const setIsProgrammaticSelectionChange = (value: boolean) => {
  isProgrammaticSelectionChange = value;
};

export const selectWordOrNextOccurrence = (editor: Editor) => {
  setIsProgrammaticSelectionChange(true);
  const allSelections = editor.listSelections();
  const { searchText, singleSearchText } = getSearchText({
    editor,
    allSelections,
    autoExpand: false,
  });

  if (searchText.length > 0 && singleSearchText) {
    const { from: latestMatchPos } = getSelectionBoundaries(
      allSelections[allSelections.length - 1],
    );
    const nextMatch = findNextMatchPosition({
      editor,
      latestMatchPos,
      searchText,
      searchWithinWords: isManualSelection,
      documentContent: editor.getValue(),
    });
    const newSelections = nextMatch
      ? allSelections.concat(nextMatch)
      : allSelections;
    editor.setSelections(newSelections);
    const lastSelection = newSelections[newSelections.length - 1];
    editor.scrollIntoView(getSelectionBoundaries(lastSelection));
  } else {
    const newSelections = [];
    for (const selection of allSelections) {
      const { from, to } = getSelectionBoundaries(selection);
      // Don't modify existing range selections
      if (from.line !== to.line || from.ch !== to.ch) {
        newSelections.push(selection);
      } else {
        newSelections.push(wordRangeAtPos(from, editor.getLine(from.line)));
        setIsManualSelection(false);
      }
    }
    editor.setSelections(newSelections);
  }
};

export const selectAllOccurrences = (editor: Editor) => {
  const allSelections = editor.listSelections();
  const { searchText, singleSearchText } = getSearchText({
    editor,
    allSelections,
    autoExpand: true,
  });
  if (!singleSearchText) {
    return;
  }
  const matches = findAllMatchPositions({
    editor,
    searchText,
    searchWithinWords: true,
    documentContent: editor.getValue(),
  });
  editor.setSelections(matches);
};

export const selectLine = (_editor: Editor, selection: EditorSelection) => {
  const { from, to } = getSelectionBoundaries(selection);
  const startOfCurrentLine = getLineStartPos(from.line);
  // if a line is already selected, expand the selection to the next line
  const startOfNextLine = getLineStartPos(to.line + 1);
  return { anchor: startOfCurrentLine, head: startOfNextLine };
};

/**
 * Finds the start position of a sentence from a given position.
 * A sentence is considered to start after a sentence-ending punctuation (. ! ?)
 * followed by zero or more closing delimiters and whitespace, or at the beginning of the line.
 */
const findSentenceStart = (
  editor: Editor,
  pos: EditorPosition,
): EditorPosition => {
  const { line } = pos;
  let ch = pos.ch;
  const lineContent = editor.getLine(line);

  // Search backwards in the current line only
  while (ch > 0) {
    // Check if the character before current position is whitespace
    if (ch > 0 && /\s/.test(lineContent.charAt(ch - 1))) {
      // Found whitespace, now look backwards to see if there's a sentence ending
      let lookbackPos = ch - 1;

      // Skip any additional whitespace
      while (lookbackPos > 0 && /\s/.test(lineContent.charAt(lookbackPos - 1))) {
        lookbackPos--;
      }

      // Skip any closing delimiters (quotes, parentheses, brackets, markdown formatting, etc.)
      while (lookbackPos > 0 && /["')\]}*_`]/.test(lineContent.charAt(lookbackPos - 1))) {
        lookbackPos--;
      }

      // Check if we have sentence-ending punctuation
      if (lookbackPos > 0 && /[.!?]/.test(lineContent.charAt(lookbackPos - 1))) {
        // Found a sentence ending! Skip forward through whitespace from current position
        while (ch < lineContent.length && /\s/.test(lineContent.charAt(ch))) {
          ch++;
        }
        return { line, ch };
      }
    }
    ch--;
  }

  // If we reached the beginning of the line, return the start of the line
  return { line, ch: 0 };
};

/**
 * Finds the end position of a sentence from a given position.
 * A sentence is considered to end at a sentence-ending punctuation (. ! ?)
 * followed by zero or more closing delimiters (quotes, parentheses, brackets, etc.)
 * or at the end of the line.
 */
const findSentenceEnd = (
  editor: Editor,
  pos: EditorPosition,
): EditorPosition => {
  const { line } = pos;
  let ch = pos.ch;
  const lineContent = editor.getLine(line);

  // Search forwards in the current line only
  while (ch < lineContent.length) {
    const char = lineContent.charAt(ch);

    // Check if we found a sentence ending (. ! ?)
    if (/[.!?]/.test(char)) {
      // Include the punctuation mark
      ch++;
      // Skip any closing delimiters (quotes, parentheses, brackets, markdown formatting, etc.)
      while (ch < lineContent.length && /["')\]}*_`]/.test(lineContent.charAt(ch))) {
        ch++;
      }
      // Skip any trailing whitespace on the same line
      while (ch < lineContent.length && /[ \t]/.test(lineContent.charAt(ch))) {
        ch++;
      }
      return { line, ch };
    }
    ch++;
  }

  // If we reached the end of the line, return the end of the line
  return { line, ch: lineContent.length };
};

export const selectSentence = (editor: Editor, selection: EditorSelection) => {
  const { from, to } = getSelectionBoundaries(selection);
  const pos = selection.head;

  // Check if we already have a selection (not just a cursor)
  const hasSelection = from.line !== to.line || from.ch !== to.ch;

  if (hasSelection) {
    // We have a selection, try to expand it
    const currentEndLine = to.line;
    const currentEndCh = to.ch;

    // Check if there's more content on the current line after the selection
    const currentLineContent = editor.getLine(currentEndLine);
    const restOfCurrentLine = currentLineContent.substring(currentEndCh);

    // If there's non-whitespace content remaining on the current line, extend to next sentence
    if (/\S/.test(restOfCurrentLine)) {
      const nextSentenceEnd = findSentenceEnd(editor, { line: currentEndLine, ch: currentEndCh });
      // Only extend if we actually found a new sentence end
      if (nextSentenceEnd.ch > currentEndCh) {
        return {
          anchor: from,
          head: nextSentenceEnd
        };
      }
    }

    // We're at the end of the current line, try to extend to the next line
    const nextLineNum = currentEndLine + 1;

    if (nextLineNum >= editor.lineCount()) {
      // No more lines to expand to
      return selection;
    }

    const nextLineContent = editor.getLine(nextLineNum);

    // If the next line is empty (or only whitespace), extend to include it
    if (nextLineContent.trim().length === 0) {
      return {
        anchor: from,
        head: { line: nextLineNum + 1, ch: 0 }
      };
    }

    // Next line has content, find the sentence end on that line
    const firstNonWhitespace = nextLineContent.search(/\S/);
    const startPos = { line: nextLineNum, ch: firstNonWhitespace >= 0 ? firstNonWhitespace : 0 };
    const nextSentenceEnd = findSentenceEnd(editor, startPos);

    return {
      anchor: from,
      head: nextSentenceEnd
    };
  }

  // No selection yet, select the current sentence (limited to current line)
  const lineContent = editor.getLine(pos.line);

  // If the current line is empty or only whitespace, select the entire line
  if (lineContent.trim().length === 0) {
    return {
      anchor: { line: pos.line, ch: 0 },
      head: { line: pos.line + 1, ch: 0 }
    };
  }

  // Skip leading whitespace to find the actual start of content
  let startSearchPos = pos;
  if (/^\s*$/.test(lineContent.substring(0, pos.ch))) {
    // If we're in leading whitespace, search from the first non-whitespace character
    const firstNonWhitespace = lineContent.search(/\S/);
    if (firstNonWhitespace !== -1) {
      startSearchPos = { line: pos.line, ch: firstNonWhitespace };
    }
  }

  const sentenceStart = findSentenceStart(editor, startSearchPos);
  const sentenceEnd = findSentenceEnd(editor, startSearchPos);

  return { anchor: sentenceStart, head: sentenceEnd };
};

export const reduceSentenceSelection = (
  editor: Editor,
  selection: EditorSelection,
) => {
  const { from, to } = getSelectionBoundaries(selection);

  // Check if we have a selection (not just a cursor)
  const hasSelection = from.line !== to.line || from.ch !== to.ch;

  if (!hasSelection) {
    // No selection to reduce
    return selection;
  }

  // If we're at the start of a line (likely after expanding through empty lines)
  if (to.ch === 0 && to.line > from.line) {
    // Go back to the end of the previous line
    const prevLine = to.line - 1;
    const prevLineContent = editor.getLine(prevLine);
    return {
      anchor: from,
      head: { line: prevLine, ch: prevLineContent.length },
    };
  }

  // Search for sentence boundaries across all lines in the selection
  // This handles cases where sentences span multiple lines
  const selectedContent = editor.getRange(from, to);
  const sentenceMatches = Array.from(selectedContent.matchAll(/[.!?](?:\s|$)/g));

  // We need at least 2 matches: one for current position, one to reduce to
  if (sentenceMatches.length >= 2) {
    // Reduce to the second-to-last sentence ending (last one is current position)
    const previousMatch = sentenceMatches[sentenceMatches.length - 2];
    const matchEnd = (previousMatch.index ?? 0) + previousMatch[0].length;

    // Convert the character offset back to line/ch position
    const newHeadOffset = editor.posToOffset(from) + matchEnd;
    const newHead = editor.offsetToPos(newHeadOffset);

    // Make sure we're not reducing to a point before or at the anchor
    if (newHead.line > from.line || (newHead.line === from.line && newHead.ch > from.ch)) {
      return {
        anchor: from,
        head: newHead,
      };
    }
  }

  // Only one sentence match found (or would reduce past anchor), return to just the first sentence
  const sentenceStart = findSentenceStart(editor, from);
  const sentenceEnd = findSentenceEnd(editor, from);
  return { anchor: sentenceStart, head: sentenceEnd };
};

/**
 * Shifts the selection to the next sentence.
 * Finds the next sentence after the current selection and moves the selection to it.
 */
export const shiftSelectionToNextSentence = (
  editor: Editor,
  selection: EditorSelection,
) => {
  const { from, to } = getSelectionBoundaries(selection);

  // Start searching from the end of the current selection
  let searchPos = to;

  // Skip any trailing whitespace at the end of current selection
  const lineContent = editor.getLine(searchPos.line);
  while (searchPos.ch < lineContent.length && /\s/.test(lineContent.charAt(searchPos.ch))) {
    searchPos = { line: searchPos.line, ch: searchPos.ch + 1 };
  }

  // If we're at the end of the line, move to the next line
  if (searchPos.ch >= lineContent.length) {
    if (searchPos.line + 1 >= editor.lineCount()) {
      // No more lines
      return selection;
    }
    searchPos = { line: searchPos.line + 1, ch: 0 };

    // Skip empty lines
    let nextLineContent = editor.getLine(searchPos.line);
    while (nextLineContent.trim().length === 0) {
      if (searchPos.line + 1 >= editor.lineCount()) {
        return selection;
      }
      searchPos = { line: searchPos.line + 1, ch: 0 };
      nextLineContent = editor.getLine(searchPos.line);
    }

    // Find first non-whitespace character
    const firstNonWhitespace = nextLineContent.search(/\S/);
    if (firstNonWhitespace !== -1) {
      searchPos = { line: searchPos.line, ch: firstNonWhitespace };
    }
  }

  // Find the start and end of the next sentence
  const nextSentenceStart = findSentenceStart(editor, searchPos);
  const nextSentenceEnd = findSentenceEnd(editor, searchPos);

  return { anchor: nextSentenceStart, head: nextSentenceEnd };
};

/**
 * Moves the current selection down by swapping with the next sentence.
 * If there's a selection, moves the entire selection (snapped to sentence boundaries).
 * If there's just a cursor, moves the sentence containing the cursor.
 * The moved selection remains selected in its new position.
 */
export const moveSentenceDown = (
  editor: Editor,
  selection: EditorSelection,
) => {
  const { from, to } = getSelectionBoundaries(selection);

  // Check if there's an actual selection (not just a cursor)
  const hasSelection = from.line !== to.line || from.ch !== to.ch;

  // Determine what to move: either the selection or the sentence at cursor
  let currentSentenceStart: EditorPosition;
  let currentSentenceEnd: EditorPosition;

  if (hasSelection) {
    // Move the entire selection, snapped to sentence boundaries
    // Start boundary: find sentence start from the beginning of selection
    currentSentenceStart = findSentenceStart(editor, from);

    // End boundary: find the sentence containing the end of selection
    const endSentenceStart = findSentenceStart(editor, to);

    // Check if findSentenceStart jumped forward past 'to' or returned exactly 'to'
    // (can happen when 'to' is in trailing whitespace or at a sentence boundary).
    // If so, 'to' is between sentences, so we should use it as-is rather than
    // expanding to include the next sentence.
    if (endSentenceStart.line > to.line ||
        (endSentenceStart.line === to.line && endSentenceStart.ch >= to.ch)) {
      // 'to' is at or after a sentence boundary, use it directly
      currentSentenceEnd = to;
    } else {
      // 'to' is within a sentence, find that sentence's end
      currentSentenceEnd = findSentenceEnd(editor, endSentenceStart);
    }
  } else {
    // No selection, just move the sentence containing the cursor
    currentSentenceStart = findSentenceStart(editor, from);
    currentSentenceEnd = findSentenceEnd(editor, from);
  }

  // Get the current sentence text
  const currentSentenceText = editor.getRange(currentSentenceStart, currentSentenceEnd);

  // Calculate current sentence length before any changes
  const currentSentenceLength = currentSentenceText.length;

  // Find the next sentence
  let nextSentenceSearchPos = currentSentenceEnd;

  // Track if we cross a paragraph break (empty line)
  let crossesParagraphBreak = false;
  let paragraphBreakStart: EditorPosition | null = null;
  let paragraphBreakEnd: EditorPosition | null = null;

  // Skip whitespace after current sentence, looking for paragraph breaks
  while (nextSentenceSearchPos.line < editor.lineCount()) {
    const lineContent = editor.getLine(nextSentenceSearchPos.line);

    if (nextSentenceSearchPos.ch < lineContent.length) {
      const char = lineContent.charAt(nextSentenceSearchPos.ch);
      if (!/\s/.test(char)) {
        break;
      }
      nextSentenceSearchPos = { line: nextSentenceSearchPos.line, ch: nextSentenceSearchPos.ch + 1 };
    } else {
      // At end of line, check if the next line is empty (paragraph break)
      if (nextSentenceSearchPos.line + 1 < editor.lineCount()) {
        const nextLine = nextSentenceSearchPos.line + 1;
        const nextLineContent = editor.getLine(nextLine);

        // If next line is empty or whitespace-only, it's a paragraph break
        if (nextLineContent.trim().length === 0) {
          crossesParagraphBreak = true;
          paragraphBreakStart = { line: nextSentenceSearchPos.line, ch: lineContent.length };
          // Find where paragraph break ends (skip all empty lines)
          let breakEndLine = nextLine;
          while (breakEndLine < editor.lineCount() && editor.getLine(breakEndLine).trim().length === 0) {
            breakEndLine++;
          }
          paragraphBreakEnd = { line: breakEndLine, ch: 0 };
          nextSentenceSearchPos = paragraphBreakEnd;
          break;
        }
      }

      // Move to next line
      if (nextSentenceSearchPos.line + 1 >= editor.lineCount()) {
        // No more content
        return selection;
      }
      nextSentenceSearchPos = { line: nextSentenceSearchPos.line + 1, ch: 0 };
    }
  }

  // If we cross a paragraph break, just move the sentence to the start of next paragraph
  // Don't swap with any sentence
  if (crossesParagraphBreak && paragraphBreakStart && paragraphBreakEnd) {
    // IMPORTANT: Calculate all offsets BEFORE modifying the document
    const paragraphBreakEndOffset = editor.posToOffset(paragraphBreakEnd);
    const currentSentenceStartOffset = editor.posToOffset(currentSentenceStart);
    const currentSentenceLength = editor.posToOffset(currentSentenceEnd) - currentSentenceStartOffset;

    // Remove sentence from current position
    editor.replaceRange('', currentSentenceStart, currentSentenceEnd);

    // Calculate insert position, adjusting for the deletion
    let insertOffset = paragraphBreakEndOffset;
    if (paragraphBreakEndOffset > currentSentenceStartOffset) {
      // The insert point is after where we deleted, so adjust backward
      insertOffset -= currentSentenceLength;
    }

    const insertPos = editor.offsetToPos(insertOffset);
    editor.replaceRange(currentSentenceText, insertPos);

    // Select the moved sentence
    const newStart = insertPos;
    const newEndOffset = editor.posToOffset(newStart) + currentSentenceLength;
    const newEnd = editor.offsetToPos(newEndOffset);

    return { anchor: newStart, head: newEnd };
  }

  // Find the boundaries of the next sentence
  const nextSentenceStart = findSentenceStart(editor, nextSentenceSearchPos);
  const nextSentenceEnd = findSentenceEnd(editor, nextSentenceSearchPos);

  // Get the next sentence text and whitespace
  const nextSentenceText = editor.getRange(nextSentenceStart, nextSentenceEnd);
  let betweenText = editor.getRange(currentSentenceEnd, nextSentenceStart);

  // Ensure there's at least one space between sentences when swapping
  // After swap: nextSentence + betweenText + currentSentence
  // So check if nextSentence (which will be first) ends with space
  if (betweenText.length === 0 && !/\s$/.test(nextSentenceText)) {
    betweenText = ' ';
  }

  // Calculate lengths before replacement
  const nextSentenceLength = nextSentenceText.length;
  const betweenLength = betweenText.length;

  // IMPORTANT: Calculate the starting offset BEFORE the replacement
  const startOffset = editor.posToOffset(currentSentenceStart);

  // Replace both sentences (swap them)
  editor.replaceRange(
    nextSentenceText + betweenText + currentSentenceText,
    currentSentenceStart,
    nextSentenceEnd
  );

  // Calculate the new position based on the string lengths
  const newStartOffset = startOffset + nextSentenceLength + betweenLength;
  const newEndOffset = newStartOffset + currentSentenceLength;

  const newStart = editor.offsetToPos(newStartOffset);
  const newEnd = editor.offsetToPos(newEndOffset);

  return { anchor: newStart, head: newEnd };
};

/**
 * Shifts the selection to the previous sentence.
 * Finds the previous sentence before the current selection and moves the selection to it.
 */
export const shiftSelectionToPreviousSentence = (
  editor: Editor,
  selection: EditorSelection,
) => {
  const { from } = getSelectionBoundaries(selection);

  // Start searching from the beginning of the current selection
  let searchPos = from;

  // Move back one character to get out of the current sentence
  if (searchPos.ch > 0) {
    searchPos = { line: searchPos.line, ch: searchPos.ch - 1 };
  } else if (searchPos.line > 0) {
    // Move to the end of the previous line
    const prevLine = searchPos.line - 1;
    const prevLineContent = editor.getLine(prevLine);
    searchPos = { line: prevLine, ch: prevLineContent.length };
  } else {
    // Already at the beginning of the document
    return selection;
  }

  // Skip backwards through any whitespace
  while (searchPos.ch > 0 || searchPos.line > 0) {
    const lineContent = editor.getLine(searchPos.line);

    if (searchPos.ch > 0) {
      const char = lineContent.charAt(searchPos.ch - 1);
      if (!/\s/.test(char)) {
        break;
      }
      searchPos = { line: searchPos.line, ch: searchPos.ch - 1 };
    } else {
      // At the beginning of a line, move to previous line
      if (searchPos.line === 0) {
        break;
      }
      const prevLine = searchPos.line - 1;
      const prevLineContent = editor.getLine(prevLine);
      searchPos = { line: prevLine, ch: prevLineContent.length };
    }
  }

  // If we're on sentence-ending punctuation or closing delimiters, move back to actual content
  // This prevents findSentenceEnd from finding the NEXT sentence
  const lineContent = editor.getLine(searchPos.line);

  // Skip backwards through any closing delimiters
  while (searchPos.ch > 0 && /["')\]}*_`]/.test(lineContent.charAt(searchPos.ch - 1))) {
    searchPos = { line: searchPos.line, ch: searchPos.ch - 1 };
  }

  // Check if we're now on sentence-ending punctuation
  if (searchPos.ch > 0 && /[.!?]/.test(lineContent.charAt(searchPos.ch - 1))) {
    // Move back before the punctuation to land in the sentence content
    if (searchPos.ch > 1) {
      searchPos = { line: searchPos.line, ch: searchPos.ch - 2 };
    }
  }

  // Find the start and end of the previous sentence
  const prevSentenceStart = findSentenceStart(editor, searchPos);
  const prevSentenceEnd = findSentenceEnd(editor, searchPos);

  return { anchor: prevSentenceStart, head: prevSentenceEnd };
};

/**
 * Moves the current selection up by swapping with the previous sentence.
 * If there's a selection, moves the entire selection (snapped to sentence boundaries).
 * If there's just a cursor, moves the sentence containing the cursor.
 * The moved selection remains selected in its new position.
 */
export const moveSentenceUp = (
  editor: Editor,
  selection: EditorSelection,
) => {
  const { from, to } = getSelectionBoundaries(selection);

  // Check if there's an actual selection (not just a cursor)
  const hasSelection = from.line !== to.line || from.ch !== to.ch;

  // Determine what to move: either the selection or the sentence at cursor
  let currentSentenceStart: EditorPosition;
  let currentSentenceEnd: EditorPosition;

  if (hasSelection) {
    // Move the entire selection, snapped to sentence boundaries
    // Start boundary: find sentence start from the beginning of selection
    currentSentenceStart = findSentenceStart(editor, from);

    // End boundary: find the sentence containing the end of selection
    const endSentenceStart = findSentenceStart(editor, to);

    // Check if findSentenceStart jumped forward past 'to' or returned exactly 'to'
    // (can happen when 'to' is in trailing whitespace or at a sentence boundary).
    // If so, 'to' is between sentences, so we should use it as-is rather than
    // expanding to include the next sentence.
    if (endSentenceStart.line > to.line ||
        (endSentenceStart.line === to.line && endSentenceStart.ch >= to.ch)) {
      // 'to' is at or after a sentence boundary, use it directly
      currentSentenceEnd = to;
    } else {
      // 'to' is within a sentence, find that sentence's end
      currentSentenceEnd = findSentenceEnd(editor, endSentenceStart);
    }
  } else {
    // No selection, just move the sentence containing the cursor
    currentSentenceStart = findSentenceStart(editor, from);
    currentSentenceEnd = findSentenceEnd(editor, from);
  }

  // Get the current sentence text
  const currentSentenceText = editor.getRange(currentSentenceStart, currentSentenceEnd);

  // Calculate current sentence length before any changes
  const currentSentenceLength = currentSentenceText.length;

  // Find the previous sentence
  let prevSentenceSearchPos = currentSentenceStart;

  // Track if we cross a paragraph break (empty line)
  let crossesParagraphBreak = false;
  let paragraphBreakStart: EditorPosition | null = null;
  let paragraphBreakEnd: EditorPosition | null = null;

  // Move back one character to get out of current sentence
  if (prevSentenceSearchPos.ch > 0) {
    prevSentenceSearchPos = { line: prevSentenceSearchPos.line, ch: prevSentenceSearchPos.ch - 1 };
  } else if (prevSentenceSearchPos.line > 0) {
    const prevLine = prevSentenceSearchPos.line - 1;
    const prevLineContent = editor.getLine(prevLine);
    prevSentenceSearchPos = { line: prevLine, ch: prevLineContent.length };
  } else {
    // Already at the beginning of the document
    return selection;
  }

  // Skip backwards through whitespace, looking for paragraph breaks
  while (prevSentenceSearchPos.ch > 0 || prevSentenceSearchPos.line > 0) {
    const lineContent = editor.getLine(prevSentenceSearchPos.line);

    if (prevSentenceSearchPos.ch > 0) {
      const char = lineContent.charAt(prevSentenceSearchPos.ch - 1);
      if (!/\s/.test(char)) {
        break;
      }
      prevSentenceSearchPos = { line: prevSentenceSearchPos.line, ch: prevSentenceSearchPos.ch - 1 };
    } else {
      // At beginning of line, check if current or previous line is empty (paragraph break)
      const currentLineContent = editor.getLine(prevSentenceSearchPos.line);
      if (currentLineContent.trim().length === 0) {
        // Found a paragraph break
        crossesParagraphBreak = true;
        // Find where paragraph break starts (skip all empty lines going back)
        let breakStartLine = prevSentenceSearchPos.line;
        while (breakStartLine > 0 && editor.getLine(breakStartLine - 1).trim().length === 0) {
          breakStartLine--;
        }
        // breakStartLine now points to first empty line after content
        paragraphBreakStart = { line: breakStartLine - 1, ch: editor.getLine(breakStartLine - 1).length };
        paragraphBreakEnd = { line: prevSentenceSearchPos.line, ch: 0 };

        // Position search at end of previous paragraph
        if (breakStartLine > 0) {
          const prevContentLine = breakStartLine - 1;
          const prevContentLineText = editor.getLine(prevContentLine);
          prevSentenceSearchPos = { line: prevContentLine, ch: prevContentLineText.length };
        }
        break;
      }

      if (prevSentenceSearchPos.line === 0) {
        break;
      }
      const prevLine = prevSentenceSearchPos.line - 1;
      const prevLineContent = editor.getLine(prevLine);
      prevSentenceSearchPos = { line: prevLine, ch: prevLineContent.length };
    }
  }

  // If we cross a paragraph break, just move the sentence to end of previous paragraph
  // Don't swap with any sentence
  if (crossesParagraphBreak && paragraphBreakStart && paragraphBreakEnd) {
    // IMPORTANT: Calculate all offsets BEFORE modifying the document
    const paragraphBreakStartOffset = editor.posToOffset(paragraphBreakStart);
    const currentSentenceStartOffset = editor.posToOffset(currentSentenceStart);
    const currentSentenceLength = editor.posToOffset(currentSentenceEnd) - currentSentenceStartOffset;

    // Remove sentence from current position
    editor.replaceRange('', currentSentenceStart, currentSentenceEnd);

    // Calculate insert position, adjusting for the deletion
    let insertOffset = paragraphBreakStartOffset;
    if (paragraphBreakStartOffset > currentSentenceStartOffset) {
      // The insert point is after where we deleted, so adjust backward
      insertOffset -= currentSentenceLength;
    }

    const insertPos = editor.offsetToPos(insertOffset);
    editor.replaceRange(currentSentenceText, insertPos);

    // Select the moved sentence
    const newStart = insertPos;
    const newEndOffset = editor.posToOffset(newStart) + currentSentenceLength;
    const newEnd = editor.offsetToPos(newEndOffset);

    return { anchor: newStart, head: newEnd };
  }

  // After skipping whitespace, we might be positioned right after the sentence-ending
  // punctuation or closing delimiters. Move back to ensure we're IN the previous sentence content.
  // This prevents findSentenceEnd from finding the CURRENT sentence's end.
  if (prevSentenceSearchPos.ch > 0) {
    const lineContent = editor.getLine(prevSentenceSearchPos.line);

    // Skip backwards through any closing delimiters
    while (prevSentenceSearchPos.ch > 0 && /["')\]}*_`]/.test(lineContent.charAt(prevSentenceSearchPos.ch - 1))) {
      prevSentenceSearchPos = { line: prevSentenceSearchPos.line, ch: prevSentenceSearchPos.ch - 1 };
    }

    // Check if we're now on sentence-ending punctuation
    if (prevSentenceSearchPos.ch > 0 && /[.!?]/.test(lineContent.charAt(prevSentenceSearchPos.ch - 1))) {
      // Move back before the punctuation to land in the sentence content
      prevSentenceSearchPos = { line: prevSentenceSearchPos.line, ch: prevSentenceSearchPos.ch - 1 };
    }

    // Move back one more character to be clearly in the sentence content
    if (prevSentenceSearchPos.ch > 0) {
      prevSentenceSearchPos = { line: prevSentenceSearchPos.line, ch: prevSentenceSearchPos.ch - 1 };
    }
  }

  // Find the boundaries of the previous sentence
  const prevSentenceStart = findSentenceStart(editor, prevSentenceSearchPos);
  const prevSentenceEnd = findSentenceEnd(editor, prevSentenceSearchPos);

  // Get the previous sentence text and whitespace
  const prevSentenceText = editor.getRange(prevSentenceStart, prevSentenceEnd);
  let betweenText = editor.getRange(prevSentenceEnd, currentSentenceStart);

  // Ensure there's at least one space between sentences when swapping
  // After swap: currentSentence + betweenText + prevSentence
  // So check if currentSentence (which will be first) ends with space
  if (betweenText.length === 0 && !/\s$/.test(currentSentenceText)) {
    betweenText = ' ';
  }

  // IMPORTANT: Calculate the starting offset BEFORE the replacement
  const startOffset = editor.posToOffset(prevSentenceStart);

  // Replace both sentences (swap them)
  editor.replaceRange(
    currentSentenceText + betweenText + prevSentenceText,
    prevSentenceStart,
    currentSentenceEnd
  );

  // Calculate the new position based on the start position and string length
  const newStartOffset = startOffset;
  const newEndOffset = newStartOffset + currentSentenceLength;

  const newStart = editor.offsetToPos(newStartOffset);
  const newEnd = editor.offsetToPos(newEndOffset);

  return { anchor: newStart, head: newEnd };
};

export const selectToEndOfSentence = (
  editor: Editor,
  selection: EditorSelection,
) => {
  const pos = selection.head;
  const sentenceEnd = findSentenceEnd(editor, pos);

  return { anchor: pos, head: sentenceEnd };
};

export const selectToStartOfSentence = (
  editor: Editor,
  selection: EditorSelection,
) => {
  const pos = selection.head;
  const sentenceStart = findSentenceStart(editor, pos);

  return { anchor: pos, head: sentenceStart };
};

export const addCursorsToSelectionEnds = (
  editor: Editor,
  emulate: CODE_EDITOR = CODE_EDITOR.VSCODE,
) => {
  // Only apply the action if there is exactly one selection
  if (editor.listSelections().length !== 1) {
    return;
  }
  const selection = editor.listSelections()[0];
  const { from, to, hasTrailingNewline } = getSelectionBoundaries(selection);
  const newSelections = [];
  // Exclude line starting at trailing newline from having cursor added
  const toLine = hasTrailingNewline ? to.line - 1 : to.line;
  for (let line = from.line; line <= toLine; line++) {
    const head = line === to.line ? to : getLineEndPos(line, editor);
    let anchor: EditorPosition;
    if (emulate === CODE_EDITOR.VSCODE) {
      anchor = head;
    } else {
      anchor = line === from.line ? from : getLineStartPos(line);
    }
    newSelections.push({
      anchor,
      head,
    });
  }
  editor.setSelections(newSelections);
};

export const goToLineBoundary = (
  editor: Editor,
  selection: EditorSelection,
  boundary: 'start' | 'end',
) => {
  const { from, to } = getSelectionBoundaries(selection);
  if (boundary === 'start') {
    return { anchor: getLineStartPos(from.line) };
  } else {
    return { anchor: getLineEndPos(to.line, editor) };
  }
};

export const navigateLine = (
  editor: Editor,
  selection: EditorSelection,
  position: 'next' | 'prev' | 'first' | 'last',
) => {
  const pos = selection.head;
  let line: number;
  let ch: number;

  if (position === 'prev') {
    line = Math.max(pos.line - 1, 0);
    const endOfLine = getLineEndPos(line, editor);
    ch = Math.min(pos.ch, endOfLine.ch);
  }
  if (position === 'next') {
    line = Math.min(pos.line + 1, editor.lineCount() - 1);
    const endOfLine = getLineEndPos(line, editor);
    ch = Math.min(pos.ch, endOfLine.ch);
  }
  if (position === 'first') {
    line = 0;
    ch = 0;
  }
  if (position === 'last') {
    line = editor.lineCount() - 1;
    const endOfLine = getLineEndPos(line, editor);
    ch = endOfLine.ch;
  }

  return { anchor: { line, ch } };
};

export const moveCursor = (
  editor: Editor,
  direction: 'up' | 'down' | 'left' | 'right',
) => {
  switch (direction) {
    case 'up':
      editor.exec('goUp');
      break;
    case 'down':
      editor.exec('goDown');
      break;
    case 'left':
      editor.exec('goLeft');
      break;
    case 'right':
      editor.exec('goRight');
      break;
  }
};

export const moveWord = (editor: Editor, direction: 'left' | 'right') => {
  switch (direction) {
    case 'left':
      editor.exec('goWordLeft');
      break;
    case 'right':
      editor.exec('goWordRight');
      break;
  }
};

export const transformCase = (
  editor: Editor,
  selection: EditorSelection,
  caseType: CASE,
) => {
  let { from, to } = getSelectionBoundaries(selection);
  let selectedText = editor.getRange(from, to);

  // apply transform on word at cursor if nothing is selected
  if (selectedText.length === 0) {
    const pos = selection.head;
    const { anchor, head } = wordRangeAtPos(pos, editor.getLine(pos.line));
    [from, to] = [anchor, head];
    selectedText = editor.getRange(anchor, head);
  }

  let replacementText = selectedText;

  switch (caseType) {
    case CASE.UPPER: {
      replacementText = selectedText.toUpperCase();
      break;
    }
    case CASE.LOWER: {
      replacementText = selectedText.toLowerCase();
      break;
    }
    case CASE.TITLE: {
      replacementText = toTitleCase(selectedText);
      break;
    }
    case CASE.NEXT: {
      replacementText = getNextCase(selectedText);
      break;
    }
  }

  editor.replaceRange(replacementText, from, to);

  return selection;
};

const expandSelection = ({
  editor,
  selection,
  openingCharacterCheck,
  matchingCharacterMap,
}: {
  editor: Editor;
  selection: EditorSelection;
  openingCharacterCheck: CheckCharacter;
  matchingCharacterMap: MatchingCharacterMap;
}) => {
  let { anchor, head } = selection;

  // in case user selects upwards
  if (anchor.line >= head.line && anchor.ch > anchor.ch) {
    [anchor, head] = [head, anchor];
  }

  const newAnchor = findPosOfNextCharacter({
    editor,
    startPos: anchor,
    checkCharacter: openingCharacterCheck,
    searchDirection: SEARCH_DIRECTION.BACKWARD,
  });
  if (!newAnchor) {
    return selection;
  }

  const newHead = findPosOfNextCharacter({
    editor,
    startPos: head,
    checkCharacter: (char: string) =>
      char === matchingCharacterMap[newAnchor.match],
    searchDirection: SEARCH_DIRECTION.FORWARD,
  });
  if (!newHead) {
    return selection;
  }

  return { anchor: newAnchor.pos, head: newHead.pos };
};

export const expandSelectionToBrackets = (
  editor: Editor,
  selection: EditorSelection,
) =>
  expandSelection({
    editor,
    selection,
    openingCharacterCheck: (char: string) => /[([{]/.test(char),
    matchingCharacterMap: MATCHING_BRACKETS,
  });

export const expandSelectionToQuotes = (
  editor: Editor,
  selection: EditorSelection,
) =>
  expandSelection({
    editor,
    selection,
    openingCharacterCheck: (char: string) => /['"`]/.test(char),
    matchingCharacterMap: MATCHING_QUOTES,
  });

export const expandSelectionToQuotesOrBrackets = (editor: Editor) => {
  const selections = editor.listSelections();
  const newSelection = expandSelection({
    editor,
    selection: selections[0],
    openingCharacterCheck: (char: string) => /['"`([{]/.test(char),
    matchingCharacterMap: MATCHING_QUOTES_BRACKETS,
  });
  editor.setSelections([...selections, newSelection]);
};

const insertCursor = (editor: Editor, lineOffset: number) => {
  const selections = editor.listSelections();
  const newSelections = [];
  for (const selection of selections) {
    const { line, ch } = selection.head;
    if (
      (line === 0 && lineOffset < 0) ||
      (line === editor.lastLine() && lineOffset > 0)
    ) {
      break;
    }
    const targetLineLength = editor.getLine(line + lineOffset).length;
    newSelections.push({
      anchor: {
        line: selection.anchor.line + lineOffset,
        ch: Math.min(selection.anchor.ch, targetLineLength),
      },
      head: {
        line: line + lineOffset,
        ch: Math.min(ch, targetLineLength),
      },
    });
  }
  editor.setSelections([...editor.listSelections(), ...newSelections]);
};

export const insertCursorAbove = (editor: Editor) => insertCursor(editor, -1);

export const insertCursorBelow = (editor: Editor) => insertCursor(editor, 1);

export const goToHeading = (
  app: App,
  editor: Editor,
  boundary: 'prev' | 'next',
) => {
  const file = app.metadataCache.getFileCache(app.workspace.getActiveFile());
  if (!file.headings || file.headings.length === 0) {
    return;
  }

  const { line } = editor.getCursor('from');
  let prevHeadingLine = 0;
  let nextHeadingLine = editor.lastLine();

  file.headings.forEach(({ position }) => {
    const { end: headingPos } = position;
    if (line > headingPos.line && headingPos.line > prevHeadingLine) {
      prevHeadingLine = headingPos.line;
    }
    if (line < headingPos.line && headingPos.line < nextHeadingLine) {
      nextHeadingLine = headingPos.line;
    }
  });

  editor.setSelection(
    boundary === 'prev'
      ? getLineEndPos(prevHeadingLine, editor)
      : getLineEndPos(nextHeadingLine, editor),
  );
};
