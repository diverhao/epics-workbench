const vscode = require("vscode");

const SEQUENCER_LANGUAGE_ID = "sequencer";
const SEQUENCER_KEYWORDS = new Set([
  "assign",
  "entry",
  "exit",
  "monitor",
  "option",
  "program",
  "ss",
  "state",
  "sync",
  "syncQ",
  "syncq",
  "to",
  "when",
]);
const SEQUENCER_TYPEWORDS = new Set([
  "char",
  "const",
  "double",
  "enum",
  "evflag",
  "float",
  "foreign",
  "int",
  "int8_t",
  "int16_t",
  "int32_t",
  "long",
  "short",
  "signed",
  "sizeof",
  "string",
  "struct",
  "typename",
  "uint8_t",
  "uint16_t",
  "uint32_t",
  "union",
  "unsigned",
  "void",
  "volatile",
  "static",
  "register",
  "auto",
  "extern",
]);
const SEQUENCER_BUILTIN_FUNCTIONS = new Set([
  "delay",
  "efClear",
  "efSet",
  "efTest",
  "efTestAndClear",
  "macValueGet",
  "optGet",
  "pvAssign",
  "pvAssignCount",
  "pvAssignSubst",
  "pvAssigned",
  "pvChannelCount",
  "pvConnectCount",
  "pvConnected",
  "pvArrayConnected",
  "pvCount",
  "pvFlush",
  "pvFlushQ",
  "pvFreeQ",
  "pvGet",
  "pvGetCancel",
  "pvArrayGetCancel",
  "pvGetComplete",
  "pvArrayGetComplete",
  "pvGetQ",
  "pvIndex",
  "pvMessage",
  "pvMonitor",
  "pvArrayMonitor",
  "pvName",
  "pvPut",
  "pvPutCancel",
  "pvArrayPutCancel",
  "pvPutComplete",
  "pvArrayPutComplete",
  "pvSeverity",
  "pvStatus",
  "pvStopMonitor",
  "pvArrayStopMonitor",
  "pvSync",
  "pvArraySync",
  "pvTimeStamp",
]);
const SEQUENCER_BUILTIN_CONSTANTS = new Set([
  "ASYNC",
  "DEFAULT",
  "DEFAULT_TIMEOUT",
  "FALSE",
  "NOEVFLAG",
  "NULL",
  "SYNC",
  "TRUE",
  "pvStatOK",
  "pvStatERROR",
  "pvStatDISCONN",
  "pvStatREAD",
  "pvStatWRITE",
  "pvStatHIHI",
  "pvStatHIGH",
  "pvStatLOLO",
  "pvStatLOW",
  "pvStatSTATE",
  "pvStatCOS",
  "pvStatCOMM",
  "pvStatTIMEOUT",
  "pvStatHW_LIMIT",
  "pvStatCALC",
  "pvStatSCAN",
  "pvStatLINK",
  "pvStatSOFT",
  "pvStatBAD_SUB",
  "pvStatUDF",
  "pvStatDISABLE",
  "pvStatSIMM",
  "pvStatREAD_ACCESS",
  "pvStatWRITE_ACCESS",
  "pvSevrOK",
  "pvSevrERROR",
  "pvSevrNONE",
  "pvSevrMINOR",
  "pvSevrMAJOR",
  "pvSevrINVALID",
  "seqg_var",
  "seqg_env",
]);
const SEQUENCER_IDENTIFIER_REGEX = /\b[A-Za-z_][A-Za-z0-9_]*\b/g;
const SEQUENCER_VARIABLE_DECLARATION_PREFIX_REGEX =
  /^\s*(?:(?:const|signed|unsigned|volatile|static|register|auto|extern)\s+)*(?:(?:enum|foreign|struct|typename|union)\s+[A-Za-z_][A-Za-z0-9_]*\s*|(?:char|double|evflag|float|int|int8_t|int16_t|int32_t|long|short|string|uint8_t|uint16_t|uint32_t|void)\s*)/;

function getSequencerDefinitionLocation(document, position) {
  if (!isSequencerDocument(document)) {
    return undefined;
  }

  const analysis = analyzeSequencerDocument(document.getText());
  const occurrence = getSequencerOccurrenceAtOffset(
    analysis,
    document.offsetAt(position),
  );
  if (!occurrence) {
    return undefined;
  }

  return createSequencerLocation(document, occurrence.definition);
}

function getSequencerReferenceLocations(document, position, includeDeclaration) {
  if (!isSequencerDocument(document)) {
    return [];
  }

  const analysis = analyzeSequencerDocument(document.getText());
  const occurrence = getSequencerOccurrenceAtOffset(
    analysis,
    document.offsetAt(position),
  );
  if (!occurrence) {
    return [];
  }

  const locations = [];
  if (includeDeclaration) {
    locations.push(createSequencerLocation(document, occurrence.definition));
  }

  for (const reference of occurrence.references) {
    locations.push(
      new vscode.Location(
        document.uri,
        new vscode.Range(
          document.positionAt(reference.start),
          document.positionAt(reference.end),
        ),
      ),
    );
  }

  return dedupeLocations(locations);
}

function getSequencerSymbolHover(document, position) {
  if (!isSequencerDocument(document)) {
    return undefined;
  }

  const analysis = analyzeSequencerDocument(document.getText());
  const occurrence = getSequencerOccurrenceAtOffset(
    analysis,
    document.offsetAt(position),
  );
  if (!occurrence?.definition) {
    return undefined;
  }

  const definition = occurrence.definition;
  const range = new vscode.Range(
    document.positionAt(occurrence.start),
    document.positionAt(occurrence.end),
  );
  if (definition.kind === "variable") {
    return createSequencerVariableSymbolHover(definition, range);
  }

  if (definition.kind === "state") {
    return createSequencerStateSymbolHover(definition, range);
  }

  return undefined;
}

function analyzeSequencerDocument(text) {
  const maskedText = maskSequencerText(text);
  const { blocks, blockByStart } = buildSequencerBlocks(maskedText);
  const definitions = collectSequencerDefinitions(maskedText, blocks, blockByStart);
  enrichSequencerDefinitions(text, maskedText, definitions);
  const definitionById = new Map(definitions.map((definition) => [definition.id, definition]));
  const occurrences = definitions.map((definition) => ({
    start: definition.nameStart,
    end: definition.nameEnd,
    definitionId: definition.id,
    isDefinition: true,
  }));

  occurrences.push(
    ...collectSequencerStateTransitionReferences(maskedText, definitions),
  );
  occurrences.push(
    ...collectSequencerVariableReferences(maskedText, definitions, occurrences),
  );

  const referencesByDefinitionId = new Map();
  for (const occurrence of occurrences) {
    if (occurrence.isDefinition) {
      continue;
    }

    const references = getOrCreateArray(
      referencesByDefinitionId,
      occurrence.definitionId,
    );
    references.push(occurrence);
  }

  return {
    definitions,
    definitionById,
    occurrences: occurrences
      .map((occurrence) => ({
        ...occurrence,
        definition: definitionById.get(occurrence.definitionId),
        references: referencesByDefinitionId.get(occurrence.definitionId) || [],
      }))
      .sort((left, right) => left.start - right.start || left.end - right.end),
  };
}

function getSequencerOccurrenceAtOffset(analysis, offset) {
  return analysis.occurrences.find(
    (occurrence) => offset >= occurrence.start && offset < occurrence.end,
  );
}

function createSequencerLocation(document, definition) {
  return new vscode.Location(
    document.uri,
    new vscode.Range(
      document.positionAt(definition.nameStart),
      document.positionAt(definition.nameEnd),
    ),
  );
}

function collectSequencerDefinitions(maskedText, blocks, blockByStart) {
  const definitions = [];
  let nextId = 1;
  const rootBlock = blocks[0];

  for (const match of maskedText.matchAll(/\bprogram\s+([A-Za-z_][A-Za-z0-9_]*)/g)) {
    const name = match[1];
    const nameStart = match.index + match[0].lastIndexOf(name);
    definitions.push({
      id: `seq-def-${nextId++}`,
      kind: "program",
      name,
      nameStart,
      nameEnd: nameStart + name.length,
      scopeStart: rootBlock.start,
      scopeEnd: rootBlock.end,
      scopeDepth: rootBlock.depth,
    });
  }

  for (const match of maskedText.matchAll(/\bss\s+([A-Za-z_][A-Za-z0-9_]*)\s*\{/g)) {
    const name = match[1];
    const nameStart = match.index + match[0].lastIndexOf(name);
    const braceIndex = match.index + match[0].lastIndexOf("{");
    const block = blockByStart.get(braceIndex) || rootBlock;
    definitions.push({
      id: `seq-def-${nextId++}`,
      kind: "stateSet",
      name,
      nameStart,
      nameEnd: nameStart + name.length,
      scopeStart: block.start,
      scopeEnd: block.end,
      scopeDepth: block.depth,
    });
  }

  for (const match of maskedText.matchAll(/\bstate\s+([A-Za-z_][A-Za-z0-9_]*)\s*\{/g)) {
    const name = match[1];
    const nameStart = match.index + match[0].lastIndexOf(name);
    const braceIndex = match.index + match[0].lastIndexOf("{");
    const block = blockByStart.get(braceIndex) || rootBlock;
    const stateSet = findSequencerContainingStateSet(definitions, nameStart);
    definitions.push({
      id: `seq-def-${nextId++}`,
      kind: "state",
      name,
      nameStart,
      nameEnd: nameStart + name.length,
      scopeStart: block.start,
      scopeEnd: block.end,
      scopeDepth: block.depth,
      stateSetId: stateSet?.id,
    });
  }

  for (const definition of collectSequencerVariableDefinitions(maskedText, blocks)) {
    definitions.push({
      ...definition,
      id: `seq-def-${nextId++}`,
    });
  }

  return definitions;
}

function buildSequencerBlocks(maskedText) {
  const rootBlock = {
    id: "seq-block-root",
    start: 0,
    end: maskedText.length,
    depth: 0,
  };
  const blocks = [rootBlock];
  const blockByStart = new Map();
  const stack = [rootBlock];

  for (let index = 0; index < maskedText.length; index += 1) {
    const character = maskedText[index];
    if (character === "{") {
      const parent = stack[stack.length - 1];
      const block = {
        id: `seq-block-${blocks.length}`,
        start: index,
        end: maskedText.length,
        depth: parent.depth + 1,
      };
      blocks.push(block);
      blockByStart.set(index, block);
      stack.push(block);
      continue;
    }

    if (character === "}" && stack.length > 1) {
      const block = stack.pop();
      block.end = index + 1;
    }
  }

  return { blocks, blockByStart };
}

function collectSequencerVariableDefinitions(maskedText, blocks) {
  const definitions = [];
  let statementStart = 0;
  let parenthesisDepth = 0;
  let bracketDepth = 0;

  for (let index = 0; index < maskedText.length; index += 1) {
    const character = maskedText[index];
    if (character === "(") {
      parenthesisDepth += 1;
      continue;
    }

    if (character === ")") {
      parenthesisDepth = Math.max(0, parenthesisDepth - 1);
      continue;
    }

    if (character === "[") {
      bracketDepth += 1;
      continue;
    }

    if (character === "]") {
      bracketDepth = Math.max(0, bracketDepth - 1);
      continue;
    }

    if (character !== ";" || parenthesisDepth !== 0 || bracketDepth !== 0) {
      continue;
    }

    const statementText = maskedText.slice(statementStart, index + 1);
    definitions.push(
      ...parseSequencerVariableDefinitionStatement(
        statementText,
        statementStart,
        blocks,
      ),
    );
    statementStart = index + 1;
  }

  return definitions;
}

function parseSequencerVariableDefinitionStatement(statementText, statementStart, blocks) {
  const declarationStart = findSequencerVariableDeclarationStart(statementText);
  if (declarationStart < 0) {
    return [];
  }

  const declarationText = statementText.slice(declarationStart);
  const prefixMatch = declarationText.match(SEQUENCER_VARIABLE_DECLARATION_PREFIX_REGEX);
  if (!prefixMatch) {
    return [];
  }
  const typeName = normalizeSequencerDeclarationType(prefixMatch[0]);

  const block = findInnermostSequencerBlock(
    blocks,
    statementStart + declarationStart + prefixMatch[0].length,
  ) || blocks[0];
  const declaratorText = declarationText
    .slice(prefixMatch[0].length)
    .replace(/;\s*$/, "");
  const definitions = [];

  for (const declarator of splitSequencerTopLevelDeclarators(
    declaratorText,
    statementStart + declarationStart + prefixMatch[0].length,
  )) {
    const assignmentIndex = getSequencerTopLevelAssignmentIndex(declarator.text);
    const declaratorValue =
      assignmentIndex >= 0
        ? declarator.text.slice(0, assignmentIndex)
        : declarator.text;
    const nameMatch = extractSequencerDeclaratorName(declaratorValue);
    if (!nameMatch) {
      continue;
    }

    definitions.push({
      kind: "variable",
      name: nameMatch.name,
      nameStart: declarator.start + nameMatch.start,
      nameEnd: declarator.start + nameMatch.end,
      typeName,
      scopeStart: block.start,
      scopeEnd: block.end,
      scopeDepth: block.depth,
    });
  }

  return definitions;
}

function findSequencerVariableDeclarationStart(statementText) {
  const candidateOffsets = [0];
  for (let index = 0; index < statementText.length; index += 1) {
    if (statementText[index] === "\n") {
      candidateOffsets.push(index + 1);
    }
  }

  for (const offset of candidateOffsets) {
    if (
      SEQUENCER_VARIABLE_DECLARATION_PREFIX_REGEX.test(
        statementText.slice(offset),
      )
    ) {
      return offset;
    }
  }

  return -1;
}

function normalizeSequencerDeclarationType(prefixText) {
  return String(prefixText || "")
    .replace(/\s+/g, " ")
    .trim();
}

function splitSequencerTopLevelDeclarators(text, baseOffset) {
  const parts = [];
  let start = 0;
  let parenthesisDepth = 0;
  let bracketDepth = 0;
  let braceDepth = 0;

  for (let index = 0; index < text.length; index += 1) {
    const character = text[index];
    if (character === "(") {
      parenthesisDepth += 1;
      continue;
    }

    if (character === ")") {
      parenthesisDepth = Math.max(0, parenthesisDepth - 1);
      continue;
    }

    if (character === "[") {
      bracketDepth += 1;
      continue;
    }

    if (character === "]") {
      bracketDepth = Math.max(0, bracketDepth - 1);
      continue;
    }

    if (character === "{") {
      braceDepth += 1;
      continue;
    }

    if (character === "}") {
      braceDepth = Math.max(0, braceDepth - 1);
      continue;
    }

    if (
      character === "," &&
      parenthesisDepth === 0 &&
      bracketDepth === 0 &&
      braceDepth === 0
    ) {
      parts.push({
        text: text.slice(start, index),
        start: baseOffset + start,
      });
      start = index + 1;
    }
  }

  if (start < text.length) {
    parts.push({
      text: text.slice(start),
      start: baseOffset + start,
    });
  }

  return parts;
}

function getSequencerTopLevelAssignmentIndex(text) {
  let parenthesisDepth = 0;
  let bracketDepth = 0;
  let braceDepth = 0;

  for (let index = 0; index < text.length; index += 1) {
    const character = text[index];
    if (character === "(") {
      parenthesisDepth += 1;
      continue;
    }

    if (character === ")") {
      parenthesisDepth = Math.max(0, parenthesisDepth - 1);
      continue;
    }

    if (character === "[") {
      bracketDepth += 1;
      continue;
    }

    if (character === "]") {
      bracketDepth = Math.max(0, bracketDepth - 1);
      continue;
    }

    if (character === "{") {
      braceDepth += 1;
      continue;
    }

    if (character === "}") {
      braceDepth = Math.max(0, braceDepth - 1);
      continue;
    }

    if (
      character === "=" &&
      parenthesisDepth === 0 &&
      bracketDepth === 0 &&
      braceDepth === 0
    ) {
      return index;
    }
  }

  return -1;
}

function extractSequencerDeclaratorName(declaratorText) {
  const matches = declaratorText.matchAll(SEQUENCER_IDENTIFIER_REGEX);
  for (const match of matches) {
    const name = match[0];
    if (isIgnoredSequencerIdentifier(name)) {
      continue;
    }

    return {
      name,
      start: match.index,
      end: match.index + name.length,
    };
  }

  return undefined;
}

function collectSequencerStateTransitionReferences(maskedText, definitions) {
  const references = [];

  for (const match of maskedText.matchAll(/\}\s*state\s+([A-Za-z_][A-Za-z0-9_]*)/g)) {
    const name = match[1];
    const start = match.index + match[0].lastIndexOf(name);
    const definition = resolveSequencerStateDefinition(definitions, name, start);
    if (!definition) {
      continue;
    }

    references.push({
      start,
      end: start + name.length,
      definitionId: definition.id,
      isDefinition: false,
    });
  }

  return references;
}

function enrichSequencerDefinitions(text, maskedText, definitions) {
  for (const definition of definitions) {
    if (definition.kind === "state") {
      definition.typeName = "state";
      definition.leadingComment = collectSequencerLeadingComment(
        text,
        definition.nameStart,
      );
      definition.whenSummaries = collectSequencerStateWhenSummaries(
        text,
        maskedText,
        definition,
      );
      continue;
    }

    if (definition.kind === "stateSet") {
      definition.typeName = "state set";
      continue;
    }

    if (definition.kind === "program") {
      definition.typeName = "program";
    }
  }
}

function collectSequencerLeadingComment(text, offset) {
  const declarationLineStart = getLineStartOffset(text, offset);
  if (declarationLineStart <= 0) {
    return undefined;
  }

  const lines = text.split("\n");
  const declarationLineIndex = getLineIndexAtOffset(text, declarationLineStart);
  let candidateLineIndex = declarationLineIndex - 1;
  while (candidateLineIndex >= 0 && !lines[candidateLineIndex].trim()) {
    candidateLineIndex -= 1;
  }

  if (candidateLineIndex < 0) {
    return undefined;
  }

  const candidateTrimmedLine = lines[candidateLineIndex].trim();
  if (!isSequencerCommentLine(candidateTrimmedLine)) {
    return undefined;
  }

  if (isSequencerLineCommentLine(candidateTrimmedLine)) {
    const collectedLines = [];
    for (let lineIndex = candidateLineIndex; lineIndex >= 0; lineIndex -= 1) {
      const trimmedLine = lines[lineIndex].trim();
      if (!isSequencerLineCommentLine(trimmedLine)) {
        break;
      }
      collectedLines.push(lines[lineIndex]);
    }
    return collectedLines.reverse().join("\n");
  }

  const collectedLines = [];
  let sawBlockStart = false;
  for (let lineIndex = candidateLineIndex; lineIndex >= 0; lineIndex -= 1) {
    const trimmedLine = lines[lineIndex].trim();
    collectedLines.push(lines[lineIndex]);
    if (trimmedLine.includes("/*")) {
      sawBlockStart = true;
      break;
    }
  }

  if (!collectedLines.length || !sawBlockStart) {
    return undefined;
  }

  return collectedLines.reverse().join("\n");
}

function getLineStartOffset(text, offset) {
  let index = Math.max(0, Math.min(offset, text.length));
  while (index > 0 && text[index - 1] !== "\n") {
    index -= 1;
  }
  return index;
}

function getLineIndexAtOffset(text, offset) {
  let lineIndex = 0;
  for (let index = 0; index < offset && index < text.length; index += 1) {
    if (text[index] === "\n") {
      lineIndex += 1;
    }
  }
  return lineIndex;
}

function isSequencerCommentLine(trimmedLine) {
  return (
    isSequencerLineCommentLine(trimmedLine) ||
    /^\/\*/.test(trimmedLine) ||
    /^\*/.test(trimmedLine) ||
    /\*\/$/.test(trimmedLine)
  );
}

function isSequencerLineCommentLine(trimmedLine) {
  return /^\/\//.test(trimmedLine);
}

function collectSequencerStateWhenSummaries(text, maskedText, stateDefinition) {
  const summaries = [];
  const stateBodyStart = Math.max(0, stateDefinition.scopeStart + 1);
  const stateBodyEnd = Math.max(stateBodyStart, stateDefinition.scopeEnd - 1);
  const stateBody = maskedText.slice(stateBodyStart, stateBodyEnd);

  for (const match of stateBody.matchAll(/\bwhen\s*\(/g)) {
    const whenStart = stateBodyStart + match.index;
    const openParenthesisIndex = whenStart + match[0].lastIndexOf("(");
    const closeParenthesisIndex = findMatchingSequencerDelimiter(
      maskedText,
      openParenthesisIndex,
      "(",
      ")",
      stateBodyEnd,
    );
    if (closeParenthesisIndex < 0) {
      continue;
    }

    const actionBlockStart = findNextSequencerNonWhitespaceIndex(
      maskedText,
      closeParenthesisIndex + 1,
      stateBodyEnd,
    );
    if (actionBlockStart < 0 || maskedText[actionBlockStart] !== "{") {
      continue;
    }

    const actionBlockEnd = findMatchingSequencerDelimiter(
      maskedText,
      actionBlockStart,
      "{",
      "}",
      stateBodyEnd,
    );
    if (actionBlockEnd < 0) {
      continue;
    }

    const targetStateStart = findNextSequencerNonWhitespaceIndex(
      maskedText,
      actionBlockEnd + 1,
      stateBodyEnd,
    );
    if (targetStateStart < 0) {
      continue;
    }

    const targetMatch = maskedText
      .slice(targetStateStart, stateBodyEnd)
      .match(/^state\s+([A-Za-z_][A-Za-z0-9_]*)/);
    if (!targetMatch) {
      continue;
    }

    const condition = text
      .slice(openParenthesisIndex + 1, closeParenthesisIndex)
      .trim();
    summaries.push({
      condition,
      targetStateName: targetMatch[1],
    });
  }

  return summaries;
}

function findMatchingSequencerDelimiter(
  text,
  startIndex,
  openCharacter,
  closeCharacter,
  endIndex,
) {
  if (startIndex < 0 || text[startIndex] !== openCharacter) {
    return -1;
  }

  let depth = 0;
  for (let index = startIndex; index < endIndex; index += 1) {
    const character = text[index];
    if (character === openCharacter) {
      depth += 1;
      continue;
    }

    if (character === closeCharacter) {
      depth -= 1;
      if (depth === 0) {
        return index;
      }
    }
  }

  return -1;
}

function findNextSequencerNonWhitespaceIndex(text, startIndex, endIndex) {
  for (let index = startIndex; index < endIndex; index += 1) {
    if (!/\s/.test(text[index])) {
      return index;
    }
  }

  return -1;
}

function collectSequencerVariableReferences(maskedText, definitions, existingOccurrences) {
  const references = [];

  for (const match of maskedText.matchAll(SEQUENCER_IDENTIFIER_REGEX)) {
    const name = match[0];
    const start = match.index;
    const end = start + name.length;

    if (
      isIgnoredSequencerIdentifier(name) ||
      existingOccurrences.some(
        (occurrence) => occurrence.start === start && occurrence.end === end,
      )
    ) {
      continue;
    }

    const definition = resolveSequencerVariableDefinition(definitions, name, start);
    if (!definition) {
      continue;
    }

    references.push({
      start,
      end,
      definitionId: definition.id,
      isDefinition: false,
    });
  }

  return references;
}

function resolveSequencerStateDefinition(definitions, name, offset) {
  const currentStateSet = findSequencerContainingStateSet(definitions, offset);
  const candidates = definitions.filter(
    (definition) =>
      definition.kind === "state" &&
      definition.name === name &&
      (!currentStateSet || definition.stateSetId === currentStateSet.id),
  );

  return candidates.sort(
    (left, right) => left.nameStart - right.nameStart,
  )[0];
}

function resolveSequencerVariableDefinition(definitions, name, offset) {
  const candidates = definitions.filter(
    (definition) =>
      definition.kind === "variable" &&
      definition.name === name &&
      definition.nameStart <= offset &&
      definition.scopeStart <= offset &&
      offset <= definition.scopeEnd,
  );

  candidates.sort((left, right) => {
    if (left.scopeDepth !== right.scopeDepth) {
      return right.scopeDepth - left.scopeDepth;
    }

    return right.nameStart - left.nameStart;
  });

  return candidates[0];
}

function findSequencerContainingStateSet(definitions, offset) {
  const candidates = definitions.filter(
    (definition) =>
      definition.kind === "stateSet" &&
      definition.scopeStart <= offset &&
      offset <= definition.scopeEnd,
  );

  candidates.sort((left, right) => {
    const leftSize = left.scopeEnd - left.scopeStart;
    const rightSize = right.scopeEnd - right.scopeStart;
    return leftSize - rightSize;
  });

  return candidates[0];
}

function findInnermostSequencerBlock(blocks, offset) {
  const candidates = blocks.filter(
    (block) => block.start <= offset && offset <= block.end,
  );

  candidates.sort((left, right) => right.depth - left.depth);
  return candidates[0];
}

function isIgnoredSequencerIdentifier(name) {
  return (
    SEQUENCER_KEYWORDS.has(name) ||
    SEQUENCER_TYPEWORDS.has(name) ||
    SEQUENCER_BUILTIN_FUNCTIONS.has(name) ||
    SEQUENCER_BUILTIN_CONSTANTS.has(name)
  );
}

function maskSequencerText(text) {
  const characters = text.split("");
  let index = 0;

  while (index < text.length) {
    if (isSequencerEmbeddedCLineStart(text, index)) {
      let end = index;
      while (end < text.length && text[end] !== "\n") {
        end += 1;
      }
      maskSequencerRange(characters, index, end);
      index = end;
      continue;
    }

    if (text.startsWith("%{", index)) {
      const endMarkerIndex = text.indexOf("}%", index + 2);
      const end = endMarkerIndex >= 0 ? endMarkerIndex + 2 : text.length;
      maskSequencerRange(characters, index, end);
      index = end;
      continue;
    }

    if (text.startsWith("//", index)) {
      let end = index;
      while (end < text.length && text[end] !== "\n") {
        end += 1;
      }
      maskSequencerRange(characters, index, end);
      index = end;
      continue;
    }

    if (text.startsWith("/*", index)) {
      const endMarkerIndex = text.indexOf("*/", index + 2);
      const end = endMarkerIndex >= 0 ? endMarkerIndex + 2 : text.length;
      maskSequencerRange(characters, index, end);
      index = end;
      continue;
    }

    if (text[index] === "\"" || text[index] === "'") {
      const quote = text[index];
      let end = index + 1;
      let escaped = false;
      while (end < text.length) {
        const character = text[end];
        if (!escaped && character === quote) {
          end += 1;
          break;
        }

        if (!escaped && character === "\n") {
          break;
        }

        escaped = !escaped && character === "\\";
        if (character !== "\\") {
          escaped = false;
        }
        end += 1;
      }

      maskSequencerRange(characters, index, end);
      index = end;
      continue;
    }

    index += 1;
  }

  return characters.join("");
}

function maskSequencerRange(characters, start, end) {
  for (let index = start; index < end; index += 1) {
    if (characters[index] !== "\n") {
      characters[index] = " ";
    }
  }
}

function isSequencerEmbeddedCLineStart(text, index) {
  if (text[index] !== "%" || text[index + 1] !== "%") {
    return false;
  }

  let lineStart = index;
  while (lineStart > 0 && text[lineStart - 1] !== "\n") {
    lineStart -= 1;
  }

  for (let cursor = lineStart; cursor < index; cursor += 1) {
    if (text[cursor] !== " " && text[cursor] !== "\t") {
      return false;
    }
  }

  return true;
}

function dedupeLocations(locations) {
  const seen = new Set();
  return locations.filter((location) => {
    const key = [
      location.uri.toString(),
      location.range.start.line,
      location.range.start.character,
      location.range.end.line,
      location.range.end.character,
    ].join(":");
    if (seen.has(key)) {
      return false;
    }

    seen.add(key);
    return true;
  });
}

function getOrCreateArray(map, key) {
  if (!map.has(key)) {
    map.set(key, []);
  }

  return map.get(key);
}

function createSequencerVariableSymbolHover(definition, range) {
  const markdown = new vscode.MarkdownString();
  markdown.appendMarkdown(`**${escapeInlineCode(definition.name)}**`);
  markdown.appendMarkdown(
    `\n\nType: \`${escapeInlineCode(definition.typeName || "variable")}\``,
  );
  return new vscode.Hover(markdown, range);
}

function createSequencerStateSymbolHover(definition, range) {
  const markdown = new vscode.MarkdownString();
  markdown.appendMarkdown(`**${escapeInlineCode(definition.name)}**`);
  markdown.appendMarkdown(
    `\n\nType: \`${escapeInlineCode(definition.typeName || "state")}\``,
  );
  if (definition.leadingComment) {
    markdown.appendMarkdown("\n\nComment:");
    markdown.appendCodeblock(definition.leadingComment, "c");
  }

  const whenSummaries = Array.isArray(definition.whenSummaries)
    ? definition.whenSummaries
    : [];
  if (!whenSummaries.length) {
    markdown.appendMarkdown("\n\nNo `when` transitions found in this state.");
    return new vscode.Hover(markdown, range);
  }

  const summaryLines = [];
  for (const summary of whenSummaries) {
    summaryLines.push(`when (${summary.condition}) {`);
    summaryLines.push("    ...");
    summaryLines.push(`} state ${summary.targetStateName}`);
    summaryLines.push("");
  }

  if (!summaryLines[summaryLines.length - 1]) {
    summaryLines.pop();
  }

  markdown.appendMarkdown("\n\nTransitions:");
  markdown.appendCodeblock(summaryLines.join("\n"), "c");
  return new vscode.Hover(markdown, range);
}

function formatSequencerText(text) {
  const normalizedText = String(text || "").replace(/\r\n/g, "\n");
  const hadTrailingNewline = normalizedText.endsWith("\n");
  const contentText = hadTrailingNewline
    ? normalizedText.slice(0, -1)
    : normalizedText;
  const lines = contentText ? contentText.split("\n") : [];
  const formattedLines = [];
  const indentUnit = " ".repeat(4);
  let indentLevel = 0;
  let inBlockComment = false;

  for (const line of lines) {
    const trimmedLine = line.trim();
    if (!trimmedLine) {
      formattedLines.push("");
      continue;
    }

    const logicalLines = expandSequencerLogicalLines(
      trimmedLine,
      inBlockComment,
    );
    for (const logicalLine of logicalLines) {
      const braceInfo = getSequencerBraceInfo(logicalLine, inBlockComment);
      const effectiveIndentLevel = braceInfo.startsWithClosingBrace
        ? Math.max(indentLevel - 1, 0)
        : indentLevel;
      const indentation = isSequencerColumnZeroLine(logicalLine)
        ? ""
        : indentUnit.repeat(effectiveIndentLevel);

      formattedLines.push(`${indentation}${logicalLine}`);
      indentLevel = Math.max(0, indentLevel + braceInfo.delta);
      inBlockComment = braceInfo.inBlockComment;
    }
  }

  let formattedText = formattedLines.join("\n");
  if (hadTrailingNewline) {
    formattedText += "\n";
  }

  return formattedText;
}

function expandSequencerLogicalLines(text, initialInBlockComment) {
  const splitLines = [];
  let inBlockComment = Boolean(initialInBlockComment);
  const compoundLines = splitSequencerCompoundLine(text, inBlockComment);

  for (const compoundLine of compoundLines) {
    const logicalLines = splitSequencerTrailingClosingBraces(
      compoundLine,
      inBlockComment,
    );
    for (const logicalLine of logicalLines) {
      splitLines.push(logicalLine);
      inBlockComment = getSequencerBraceInfo(
        logicalLine,
        inBlockComment,
      ).inBlockComment;
    }
  }

  return splitLines;
}

function splitSequencerCompoundLine(text, initialInBlockComment) {
  if (!text || isSequencerColumnZeroLine(text)) {
    return [text];
  }

  const parts = [];
  let segmentStart = 0;
  let inBlockComment = Boolean(initialInBlockComment);
  let inDoubleQuote = false;
  let inSingleQuote = false;
  let escaped = false;

  for (let index = 0; index < text.length; index += 1) {
    const character = text[index];
    const nextCharacter = text[index + 1];

    if (inBlockComment) {
      if (character === "*" && nextCharacter === "/") {
        inBlockComment = false;
        index += 1;
      }
      continue;
    }

    if (inDoubleQuote) {
      if (escaped) {
        escaped = false;
        continue;
      }

      if (character === "\\") {
        escaped = true;
        continue;
      }

      if (character === "\"") {
        inDoubleQuote = false;
      }
      continue;
    }

    if (inSingleQuote) {
      if (escaped) {
        escaped = false;
        continue;
      }

      if (character === "\\") {
        escaped = true;
        continue;
      }

      if (character === "'") {
        inSingleQuote = false;
      }
      continue;
    }

    if (character === "/" && nextCharacter === "*") {
      inBlockComment = true;
      index += 1;
      continue;
    }

    if (character === "/" && nextCharacter === "/") {
      break;
    }

    if (character === "%" && nextCharacter === "{") {
      index += 1;
      continue;
    }

    if (character === "}" && nextCharacter === "%") {
      index += 1;
      continue;
    }

    if (character === "\"") {
      inDoubleQuote = true;
      continue;
    }

    if (character === "'") {
      inSingleQuote = true;
      continue;
    }

    if (character === "{") {
      const nextNonWhitespaceIndex = getNextNonWhitespaceIndex(text, index + 1);
      if (nextNonWhitespaceIndex >= 0) {
        const head = text.slice(segmentStart, index + 1).trim();
        if (head) {
          parts.push(head);
        }
        segmentStart = nextNonWhitespaceIndex;
      }
      continue;
    }

    if (character === "}") {
      const currentSegment = text.slice(segmentStart, index).trim();
      const nextNonWhitespaceIndex = getNextNonWhitespaceIndex(text, index + 1);
      if (currentSegment && nextNonWhitespaceIndex >= 0) {
        parts.push(currentSegment);
        segmentStart = index;
      }
    }
  }

  const tail = text.slice(segmentStart).trim();
  if (tail) {
    parts.push(tail);
  }

  return parts.length ? parts : [text];
}

function splitSequencerTrailingClosingBraces(text, initialInBlockComment) {
  if (!text || isSequencerColumnZeroLine(text)) {
    return [text];
  }

  let inBlockComment = Boolean(initialInBlockComment);
  let inDoubleQuote = false;
  let inSingleQuote = false;
  let escaped = false;
  let firstCodeTokenIndex = -1;
  let trailingBracePositions = [];

  for (let index = 0; index < text.length; index += 1) {
    const character = text[index];
    const nextCharacter = text[index + 1];

    if (inBlockComment) {
      if (character === "*" && nextCharacter === "/") {
        inBlockComment = false;
        index += 1;
      }
      continue;
    }

    if (inDoubleQuote) {
      if (escaped) {
        escaped = false;
        continue;
      }

      if (character === "\\") {
        escaped = true;
        continue;
      }

      if (character === "\"") {
        inDoubleQuote = false;
      }
      continue;
    }

    if (inSingleQuote) {
      if (escaped) {
        escaped = false;
        continue;
      }

      if (character === "\\") {
        escaped = true;
        continue;
      }

      if (character === "'") {
        inSingleQuote = false;
      }
      continue;
    }

    if (character === "/" && nextCharacter === "*") {
      inBlockComment = true;
      index += 1;
      continue;
    }

    if (character === "/" && nextCharacter === "/") {
      break;
    }

    if (character === "%" && nextCharacter === "{") {
      if (firstCodeTokenIndex < 0) {
        firstCodeTokenIndex = index;
      }
      trailingBracePositions = [];
      index += 1;
      continue;
    }

    if (character === "}" && nextCharacter === "%") {
      if (firstCodeTokenIndex < 0) {
        firstCodeTokenIndex = index;
      }
      trailingBracePositions = [];
      index += 1;
      continue;
    }

    if (character === "\"") {
      inDoubleQuote = true;
      continue;
    }

    if (character === "'") {
      inSingleQuote = true;
      continue;
    }

    if (/\s/.test(character)) {
      continue;
    }

    if (firstCodeTokenIndex < 0) {
      firstCodeTokenIndex = index;
    }

    if (character === "}") {
      trailingBracePositions.push(index);
      continue;
    }

    trailingBracePositions = [];
  }

  if (!trailingBracePositions.length) {
    return [text];
  }

  const firstTrailingBraceIndex = trailingBracePositions[0];
  if (firstCodeTokenIndex < 0 || firstCodeTokenIndex >= firstTrailingBraceIndex) {
    return [text];
  }

  const head = text.slice(0, firstTrailingBraceIndex).trimEnd();
  if (!head) {
    return [text];
  }

  return [head, ...trailingBracePositions.map(() => "}")];
}

function getNextNonWhitespaceIndex(text, startIndex) {
  for (let index = startIndex; index < text.length; index += 1) {
    if (!/\s/.test(text[index])) {
      return index;
    }
  }

  return -1;
}

function isSequencerColumnZeroLine(trimmedLine) {
  return /^#/.test(trimmedLine) || /^%%/.test(trimmedLine);
}

function getSequencerBraceInfo(text, initialInBlockComment) {
  let delta = 0;
  let inBlockComment = Boolean(initialInBlockComment);
  let inDoubleQuote = false;
  let inSingleQuote = false;
  let escaped = false;
  let sawCodeToken = false;
  let startsWithClosingBrace = false;

  for (let index = 0; index < text.length; index += 1) {
    const character = text[index];
    const nextCharacter = text[index + 1];

    if (inBlockComment) {
      if (character === "*" && nextCharacter === "/") {
        inBlockComment = false;
        index += 1;
      }
      continue;
    }

    if (inDoubleQuote) {
      if (escaped) {
        escaped = false;
        continue;
      }

      if (character === "\\") {
        escaped = true;
        continue;
      }

      if (character === "\"") {
        inDoubleQuote = false;
      }
      continue;
    }

    if (inSingleQuote) {
      if (escaped) {
        escaped = false;
        continue;
      }

      if (character === "\\") {
        escaped = true;
        continue;
      }

      if (character === "'") {
        inSingleQuote = false;
      }
      continue;
    }

    if (character === "/" && nextCharacter === "*") {
      inBlockComment = true;
      index += 1;
      continue;
    }

    if (character === "/" && nextCharacter === "/") {
      break;
    }

    if (character === "%" && nextCharacter === "{") {
      index += 1;
      continue;
    }

    if (character === "}" && nextCharacter === "%") {
      sawCodeToken = true;
      index += 1;
      continue;
    }

    if (character === "\"") {
      inDoubleQuote = true;
      continue;
    }

    if (character === "'") {
      inSingleQuote = true;
      continue;
    }

    if (/\s/.test(character)) {
      continue;
    }

    if (!sawCodeToken && character === "}") {
      startsWithClosingBrace = true;
    }
    sawCodeToken = true;

    if (character === "{") {
      delta += 1;
    } else if (character === "}") {
      delta -= 1;
    }
  }

  return {
    delta,
    inBlockComment,
    startsWithClosingBrace,
  };
}

function escapeInlineCode(value) {
  return String(value).replace(/`/g, "\\`");
}

function isSequencerDocument(document) {
  return document && document.languageId === SEQUENCER_LANGUAGE_ID;
}

module.exports = {
  formatSequencerText,
  getSequencerDefinitionLocation,
  getSequencerReferenceLocations,
  getSequencerSymbolHover,
};
