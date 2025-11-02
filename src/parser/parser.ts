// Excel Formula Parser - Pure parsing logic

import { FormulaTypes } from "./formula-types";
import { ExcelExpression, SubExpression, FormulaExpr } from "./types";

export class ExcelFormulaParser {
  private static readonly formulaTypes: Array<string> = Object.values(FormulaTypes);

  private static parseExpressions(
    parent: ExcelExpression,
    startIndex: number,
    input: string,
    separator: string,
  ): number {
    let token = "";
    let i: number = startIndex;

    while (i < input.length) {
      if (input[i] === "(") {
        if (token.length > 0) {
          // End of expression
          parent.addChild(new ExcelExpression(token));
          token = "";
        }
        const expression = new SubExpression("");
        parent.addChild(expression);
        i = this.parseExpressions(expression, i + 1, input, separator);
      } else if (input[i] === ")") {
        if (token.length > 0) {
          // End of expression
          parent.addChild(new ExcelExpression(token));
        }
        return i;
      } else {
        token += input[i];
        if (this.formulaTypes.includes(token.toUpperCase())) {
          // Parse formula and jump to next expression
          i = this.parseFormula(
            parent,
            FormulaTypes[token.toUpperCase() as keyof typeof FormulaTypes],
            i + 1,
            input,
            separator,
          );
          // Token consumed
          token = "";
        }
      }

      i++;
    }

    if (token.length > 0) {
      parent.addChild(new ExcelExpression(token));
    }

    return i;
  }

  private static parseFormula(
    parent: ExcelExpression,
    formula: FormulaTypes,
    startIndex: number,
    input: string,
    separator: string,
  ): number {
    const formulaExpr = new FormulaExpr("", formula);
    parent.addChild(formulaExpr);

    let i = this.parserAdvanceTo("(", startIndex, input);
    let token = "";
    let braceCount = 0;

    while (i < input.length) {
      if (input[i] === separator && braceCount === 0) {
        if (token.trim().length > 0) {
          const expression = new ExcelExpression(token.trim());
          this.parseExpressions(expression, 0, token.trim(), separator);

          if (expression.getChilds().length === 1) {
            formulaExpr.addChild(expression.getChilds()[0]);
          } else if (expression.getChilds().length > 1) {
            formulaExpr.addChild(expression);
          } else {
            // No children found, add the original token as a simple expression
            formulaExpr.addChild(new ExcelExpression(token.trim()));
          }
        }
        token = "";
      } else {
        if (input[i] === "(") {
          braceCount++;
        } else if (input[i] === ")") {
          braceCount--;
        }

        if (braceCount < 0) {
          if (token.trim().length > 0) {
            const expression = new ExcelExpression(token.trim());
            this.parseExpressions(expression, 0, token.trim(), separator);

            if (expression.getChilds().length === 1) {
              formulaExpr.addChild(expression.getChilds()[0]);
            } else if (expression.getChilds().length > 1) {
              formulaExpr.addChild(expression);
            } else {
              // No children found, add the original token as a simple expression
              formulaExpr.addChild(new ExcelExpression(token.trim()));
            }
          }
          return i;
        }

        token += input[i];
      }
      i++;
    }

    return i;
  }

  private static parserAdvanceTo(char: string, startIndex: number, input: string): number {
    let i = startIndex;
    while (i < input.length && input[i] !== char) {
      i++;
    }
    return i + 1; // Move past the found character
  }

  static parse(formula: string): ExcelExpression {
    // Remove leading = if present
    let cleanFormula = formula.trim();
    if (cleanFormula.startsWith("=")) {
      cleanFormula = cleanFormula.substring(1);
    }

    // Determine separator (European uses ; US uses ,)
    const separator = cleanFormula.includes(";") ? ";" : ",";

    try {
      const baseExpression = new ExcelExpression(cleanFormula);
      this.parseExpressions(baseExpression, 0, cleanFormula, separator);
      return baseExpression;
    } catch (error) {
      console.error("Parse error:", error);
      // Return a simple expression with the original formula
      return new ExcelExpression(cleanFormula);
    }
  }
}