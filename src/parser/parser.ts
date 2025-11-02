// Excel Formula Parser - Pure parsing logic

import { FormulaTypes } from "./formula-types";
import { ExcelExpression, SubExpression, FormulaExpr, OperatorExpression } from "./types";

export class ExcelFormulaParser {
  private static readonly formulaTypes: Array<string> = Object.values(FormulaTypes);
  
  // Define operators in order of precedence (longest first to avoid partial matches)
  private static readonly operators: Array<string> = [
    "<=", ">=", "<>", "!=", "==", // Comparison operators (2 chars)
    "+", "-", "*", "/", "^", "&", "=", "<", ">", ":", // Single char operators
  ];

  private static findOperatorAt(input: string, position: number): string | null {
    // Check for 2-character operators first
    if (position < input.length - 1) {
      const twoChar = input.substring(position, position + 2);
      if (this.operators.includes(twoChar)) {
        return twoChar;
      }
    }
    
    // Check for single-character operators
    const oneChar = input.charAt(position);
    if (this.operators.includes(oneChar)) {
      return oneChar;
    }
    
    return null;
  }

  private static parseExpressions(
    parent: ExcelExpression,
    startIndex: number,
    input: string,
    separator: string,
  ): number {
    let token = "";
    let i: number = startIndex;

    while (i < input.length) {
      const operator = this.findOperatorAt(input, i);
      
      if (input[i] === "(") {
        if (token.trim().length > 0) {
          // Check if token is a function name
          if (this.formulaTypes.includes(token.trim().toUpperCase())) {
            i = this.parseFormula(
              parent,
              FormulaTypes[token.trim().toUpperCase() as keyof typeof FormulaTypes],
              i,
              input,
              separator,
            );
            token = "";
          } else {
            // Regular token before parentheses
            parent.addChild(new ExcelExpression(token.trim()));
            token = "";
            const expression = new SubExpression("");
            parent.addChild(expression);
            i = this.parseExpressions(expression, i + 1, input, separator);
          }
        } else {
          // Parentheses without preceding token
          const expression = new SubExpression("");
          parent.addChild(expression);
          i = this.parseExpressions(expression, i + 1, input, separator);
        }
      } else if (input[i] === ")") {
        if (token.trim().length > 0) {
          parent.addChild(new ExcelExpression(token.trim()));
        }
        return i;
      } else if (operator) {
        // Found an operator
        if (token.trim().length > 0) {
          parent.addChild(new ExcelExpression(token.trim()));
          token = "";
        }
        parent.addChild(new OperatorExpression(operator));
        i += operator.length - 1; // Skip the operator characters (-1 because i++ will happen)
      } else if (input[i] === " " || input[i] === "\t" || input[i] === "\n") {
        // Handle whitespace - just add to token for now, we'll trim later
        token += input[i];
      } else {
        // Regular character
        token += input[i];
      }

      i++;
    }

    if (token.trim().length > 0) {
      parent.addChild(new ExcelExpression(token.trim()));
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
          const paramExpression = new ExcelExpression(token.trim());
          this.parseExpressions(paramExpression, 0, token.trim(), separator);

          if (paramExpression.getChilds().length === 1) {
            formulaExpr.addChild(paramExpression.getChilds()[0]);
          } else if (paramExpression.getChilds().length > 1) {
            formulaExpr.addChild(paramExpression);
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
          // End of function parameters
          if (token.trim().length > 0) {
            const paramExpression = new ExcelExpression(token.trim());
            this.parseExpressions(paramExpression, 0, token.trim(), separator);

            if (paramExpression.getChilds().length === 1) {
              formulaExpr.addChild(paramExpression.getChilds()[0]);
            } else if (paramExpression.getChilds().length > 1) {
              formulaExpr.addChild(paramExpression);
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