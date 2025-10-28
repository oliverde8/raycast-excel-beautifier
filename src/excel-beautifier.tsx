import { Detail, Clipboard, showToast, Toast } from "@raycast/api";
import { useEffect, useState } from "react";

// Type definitions for Excel formula tokens
type TokenType = 'function' | 'cell' | 'range' | 'string' | 'number' | 'operator' | 'identifier' | 'unknown';

interface FormulaToken {
  type: TokenType;
  value: string;
  position: number;
}

interface FormulaAnalysis {
  tokenCount: number;
  functionCount: number;
  cellReferences: number;
  maxNesting: number;
}

interface BeautificationResult {
  beautified: string;
  tokens: FormulaToken[];
  analysis: FormulaAnalysis;
}

class ExcelFormulaBeautifier {
  private static readonly FUNCTIONS = new Set([
    'SUM', 'AVERAGE', 'COUNT', 'MAX', 'MIN', 'IF', 'VLOOKUP', 'HLOOKUP', 'INDEX', 'MATCH',
    'CONCATENATE', 'LEFT', 'RIGHT', 'MID', 'LEN', 'FIND', 'SUBSTITUTE', 'UPPER', 'LOWER',
    'AND', 'OR', 'NOT', 'IFERROR', 'ISERROR', 'ISBLANK', 'SUMIF', 'SUMIFS', 'COUNTIF',
    'COUNTIFS', 'AVERAGEIF', 'AVERAGEIFS', 'ROUND', 'ROUNDUP', 'ROUNDDOWN', 'ABS', 'MOD',
    'TODAY', 'NOW', 'DATE', 'YEAR', 'MONTH', 'DAY', 'WEEKDAY', 'NETWORKDAYS', 'PMT', 'PV', 'FV'
  ]);

  // The lexer ... or yacc.
  // TODO this don't work ideally which causes the beatification not to be perfect...
  private static tokenize(formula: string): FormulaToken[] {
    const tokens: FormulaToken[] = [];
    let i = 0;

    while (i < formula.length) {
      const char = formula[i];
      const position = i;

      if (/\s/.test(char)) {
        i++;
        continue;
      }

      // Search for formula
      if (char === '"') {
        let value = '"';
        i++;
        while (i < formula.length && formula[i] !== '"') {
          value += formula[i];
          if (formula[i] === '"' && formula[i + 1] === '"') {
            value += formula[i + 1];
            i += 2;
          } else {
            i++;
          }
        }
        if (i < formula.length) value += formula[i++];
        tokens.push({ type: 'string', value, position });
        continue;
      }

      // Function names and cell references
      if (/[A-Za-z_]/.test(char)) {
        let value = '';
        while (i < formula.length && /[A-Za-z0-9_.$:]/.test(formula[i])) {
          value += formula[i++];
        }

        const upperValue = value.toUpperCase();
        if (this.FUNCTIONS.has(upperValue)) {
          tokens.push({ type: 'function', value: upperValue, position });
        } else if (/^[A-Z]+[0-9]+$/.test(value) || /^\$?[A-Z]+\$?[0-9]+$/.test(value)) {
          tokens.push({ type: 'cell', value, position });
        } else if (value.includes('!') || value.includes(':')) {
          tokens.push({ type: 'range', value, position });
        } else {
          tokens.push({ type: 'identifier', value, position });
        }
        continue;
      }
      // Numbers
      if (/[0-9.]/.test(char)) {
        let value = '';
        while (i < formula.length && /[0-9.]/.test(formula[i])) {
          value += formula[i++];
        }
        tokens.push({ type: 'number', value, position });
        continue;
      }
      // Operators and punctuation
      if ('()[]{}+*-/=<>!&,;'.includes(char)) {
        let value = char;
        i++;
        // Handle multi-character operators
        if (char === '<' && i < formula.length && formula[i] === '=') {
          value += formula[i++];
        } else if (char === '>' && i < formula.length && formula[i] === '=') {
          value += formula[i++];
        } else if (char === '<' && i < formula.length && formula[i] === '>') {
          value += formula[i++];
        }
        // Normalize semicolons to commas (European vs US format)
        if (value === ';') {
          value = ',';
        }
        tokens.push({ type: 'operator', value, position });
        continue;
      }
      // Unknown character
      tokens.push({ type: 'unknown', value: char, position });
      i++;
    }
    return tokens;
  }
  private static isClosingFunctionParenthesis(tokens: FormulaToken[], currentIndex: number): boolean {
    let depth = 0;
    for (let i = currentIndex; i >= 0; i--) {
      const token = tokens[i];
      if (token.value === ')') {
        depth++;
      } else if (token.value === '(') {
        depth--;
        if (depth === 0) {
          // Found the matching opening parenthesis, check if it's preceded by a function
          const prevToken = i > 0 ? tokens[i - 1] : null;
          return prevToken && prevToken.type === 'function';
        }
      }
    }
    return false;
  }
  static beautify(formula: string): BeautificationResult {
    // Remove leading = if present
    let cleanFormula = formula.trim();
    const hasEquals = cleanFormula.startsWith('=');
    if (hasEquals) {
      cleanFormula = cleanFormula.substring(1);
    }
    const tokens = this.tokenize(cleanFormula);
    let result = '';
    let indentLevel = 0;
    const indentSize = 2;
    let needsNewline = false;
    for (let i = 0; i < tokens.length; i++) {
      const token = tokens[i];
      const nextToken = tokens[i + 1];
      const prevToken = tokens[i - 1];
      if (needsNewline) {
        result += '\n' + ' '.repeat(indentLevel * indentSize);
        needsNewline = false;
      }
      if (token.type === 'operator') {
        switch (token.value) {
          case '(':
          {
            result += token.value;
            // Only indent and create new lines for function calls, not mathematical grouping
            const isFunctionCall = prevToken && prevToken.type === 'function';
            if (isFunctionCall) {
              indentLevel++;
              if (nextToken && nextToken.type !== 'operator') {
                needsNewline = true;
              }
            }
            break;
          }
          case ')':
            // Only dedent if we're closing a function call
          {
            const isClosingFunction = ExcelFormulaBeautifier.isClosingFunctionParenthesis(tokens, i);
            if (isClosingFunction) {
              indentLevel = Math.max(0, indentLevel - 1);
              if (prevToken && prevToken.type !== 'operator') {
                result += '\n' + ' '.repeat(indentLevel * indentSize);
              }
            }
            result += token.value;
            break;
          }
          case ',':
            result += token.value;
            if (indentLevel > 0) {
              needsNewline = true;
            } else {
              result += ' ';
            }
            break;
          case ';':
            // Semicolons should be treated like commas (already normalized in tokenizer)
            result += ',';
            if (indentLevel > 0) {
              needsNewline = true;
            } else {
              result += ' ';
            }
            break;
          default:
            // Handle spacing around operators
            {
              const needsSpaceBefore = prevToken &&
                !['(', ',', ';'].includes(prevToken.value) &&
                ['=', '+', '-', '*', '/', '<', '>', '<=', '>=', '<>', '&'].includes(token.value);
              const needsSpaceAfter = nextToken &&
                ![')', ',', ';'].includes(nextToken.value) &&
                ['=', '+', '-', '*', '/', '<', '>', '<=', '>=', '<>', '&'].includes(token.value);
              if (needsSpaceBefore) result += ' ';
              result += token.value;

              if (needsSpaceAfter) result += ' ';
              break;
            }
        }
      } else if (token.type === 'function') {
        if (prevToken && !['(', ',', ';'].includes(prevToken.value)) {
          result += ' ';
        }
        result += token.value;
      } else {
        result += token.value;
      }
    }

    const analysis: FormulaAnalysis = {
      tokenCount: tokens.length,
      functionCount: tokens.filter(t => t.type === 'function').length,
      cellReferences: tokens.filter(t => t.type === 'cell' || t.type === 'range').length,
      maxNesting: Math.max(...tokens.map((_, i, arr) => 
        arr.slice(0, i + 1).filter(t => t.value === '(').length
      ), 0)
    };

    return {
      beautified: (hasEquals ? '=' : '') + result,
      tokens,
      analysis
    };
  }
}

interface FormulaData {
  original: string;
  beautified: string;
  tokens?: FormulaToken[];
  analysis?: FormulaAnalysis;
}

export default function Command() {
  const [formulaData, setFormulaData] = useState<FormulaData | null>(null);
  const [isLoading, setIsLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    async function getClipboardContent() {
      try {
        const clipboardContent = await Clipboard.readText();
        if (!clipboardContent) {
          setError("No text found in clipboard");
          setIsLoading(false);
          return;
        }

        // Check if it looks like an Excel formula
        const trimmed = clipboardContent.trim();
        if (!trimmed.startsWith('=')) {
          // Try to beautify anyway, maybe user copied formula without =
          showToast({
            style: Toast.Style.Animated,
            title: "Warning",
            message: "Content doesn't start with '='. Treating as formula anyway."
          });
        }

        try {
          // Use our custom beautifier designed for Node.js environments
          const result = ExcelFormulaBeautifier.beautify(trimmed);
          setFormulaData({
            original: trimmed,
            beautified: result.beautified,
            tokens: result.tokens,
          });
        } catch (formulaError) {
          // If beautification fails, show error with the original content
          setError(`Invalid Excel formula: ${formulaError instanceof Error ? formulaError.message : String(formulaError)}`);
          setIsLoading(false);
          return;
        }
        setIsLoading(false);
      } catch (err) {
        setError(`Failed to read clipboard: ${err}`);
        setIsLoading(false);
      }
    }

    getClipboardContent();
  }, []);

  if (isLoading) {
    return <Detail isLoading={true} markdown="Reading from clipboard..." />;
  }

  if (error) {
    return <Detail markdown={`# Error\n\n${error}`} />;
  }

  if (!formulaData) {
    return <Detail markdown="# No formula data available" />;
  }

  const markdown = `## Original Formula
\`\`\`
${formulaData.original}
\`\`\`

## Beautified Formula
\`\`\`excel
${formulaData.beautified}
\`\`\`
`;

  return <Detail markdown={markdown} />;
}