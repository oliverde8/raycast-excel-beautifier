// Excel Formula Parser and Beautifier

import { FormulaTypes } from "./formula-types";

class ExcelExpression {
  readonly original: string;
  private childs: Array<ExcelExpression> = [];

  constructor(original: string) {
    this.original = original;
  }

  addChild(child: ExcelExpression): void {
    this.childs.push(child);
  }

  getChilds(): Array<ExcelExpression> {
    return this.childs;
  }
}

class SubExpression extends ExcelExpression {}

class FormulaExpr extends ExcelExpression {
  readonly formula: FormulaTypes;

  constructor(original: string, formula: FormulaTypes) {
    super(original);
    this.formula = formula;
  }
}

interface BeautificationResult {
  beautified: string;
}

export class ExcelFormulaBeautifier {
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

  private static prettyPrint(expression: ExcelExpression, depth: number): string {
    const children = expression.getChilds();
    
    // If no children, return the original content
    if (children.length === 0) {
      return expression.original || "";
    }

    let result = "";
    
    for (let i = 0; i < children.length; i++) {
      const child = children[i];
      
      if (child instanceof FormulaExpr) {
        result += this.formatFunction(child, depth);
      } else if (child instanceof SubExpression) {
        result += this.formatSubExpression(child, depth);
      } else {
        // For regular expressions, check if they have children or use original
        if (child.getChilds().length === 0) {
          result += child.original || "";
        } else {
          result += this.prettyPrint(child, depth);
        }
      }
      
      // Add spacing between elements if needed
      if (i < children.length - 1) {
        const nextChild = children[i + 1];
        if (this.needsSpacing(child, nextChild)) {
          result += " ";
        }
      }
    }
    
    return result;
  }

  private static formatFunction(func: FormulaExpr, depth: number): string {
    const params = func.getChilds();
    if (params.length === 0) {
      return func.formula + "()";
    }

    // Check if this function should be formatted inline
    if (this.shouldFormatInline(func)) {
      let result = func.formula + "(";
      for (let i = 0; i < params.length; i++) {
        const param = params[i];
        const isLast = i === params.length - 1;
        
        if (param instanceof FormulaExpr && this.shouldFormatInline(param)) {
          result += this.formatFunction(param, depth);
        } else if (param instanceof SubExpression) {
          result += this.formatSubExpression(param, depth);
        } else {
          const paramContent = param.getChilds().length === 0 
            ? param.original 
            : this.prettyPrint(param, depth);
          result += paramContent || "";
        }
        
        if (!isLast) {
          result += "; ";
        }
      }
      result += ")";
      return result;
    }

    // Multi-line formatting for complex functions
    let result = func.formula + "(\n";
    
    for (let i = 0; i < params.length; i++) {
      const param = params[i];
      const isLast = i === params.length - 1;
      
      result += this.indent(depth + 1);
      
      if (param instanceof FormulaExpr) {
        result += this.formatFunction(param, depth + 1);
      } else if (param instanceof SubExpression) {
        result += this.formatSubExpression(param, depth + 1);
      } else {
        // For parameter content, preserve original or format children
        const paramContent = param.getChilds().length === 0 
          ? param.original 
          : this.prettyPrint(param, depth + 1);
        result += paramContent || "";
      }
      
      if (!isLast) {
        result += ";";
      }
      result += "\n";
    }
    
    result += this.indent(depth) + ")";
    return result;
  }

  private static shouldFormatInline(func: FormulaExpr): boolean {
    const params = func.getChilds();
    
    // No parameters or only one parameter
    if (params.length === 0 || params.length === 1) {
      // Check if the single parameter is simple
      if (params.length === 1) {
        const param = params[0];
        // If it's a nested function, don't inline
        if (param instanceof FormulaExpr) {
          return false;
        }
        // If it's a complex sub-expression, don't inline
        if (param instanceof SubExpression && this.containsFunction(param)) {
          return false;
        }
        // Check the content length
        const content = param.original || this.prettyPrint(param, 0);
        return content.length <= 30;
      }
      return true;
    }
    
    // Multiple parameters - only inline if all are very simple and short
    if (params.length <= 3) {
      let totalLength = 0;
      for (const param of params) {
        if (param instanceof FormulaExpr || (param instanceof SubExpression && this.containsFunction(param))) {
          return false;
        }
        const content = param.original || this.prettyPrint(param, 0);
        totalLength += content.length;
        if (totalLength > 40) {
          return false;
        }
      }
      return totalLength <= 40;
    }
    
    return false;
  }

  private static formatSubExpression(subExpr: SubExpression, depth: number): string {
    const content = this.prettyPrint(subExpr, depth + 1);
    
    // If it's simple, keep on one line
    if (!content.includes('\n') && content.length < 50 && !this.containsFunction(subExpr)) {
      return "(" + content + ")";
    }
    
    // Complex sub-expression gets multi-line formatting
    return "(\n" + this.indent(depth + 1) + content + "\n" + this.indent(depth) + ")";
  }

  private static containsFunction(expr: ExcelExpression): boolean {
    if (expr instanceof FormulaExpr) {
      return true;
    }
    return expr.getChilds().some(child => this.containsFunction(child));
  }

  private static needsSpacing(current: ExcelExpression, next: ExcelExpression): boolean {
    const currentText = current.original?.trim();
    const nextText = next.original?.trim();
    
    if (!currentText || !nextText) return false;
    
    const operators = ['+', '-', '*', '/', '=', '<', '>', '<=', '>=', '<>', '&'];
    return operators.includes(currentText) || operators.includes(nextText);
  }

  private static indent(depth: number): string {
    return "    ".repeat(depth);
  }

  static beautify(formula: string): BeautificationResult {
    // Remove leading = if present
    let cleanFormula = formula.trim();
    const hasEquals = cleanFormula.startsWith("=");
    if (hasEquals) {
      cleanFormula = cleanFormula.substring(1);
    }

    // Determine separator (European uses ; US uses ,)
    const separator = cleanFormula.includes(";") ? ";" : ",";

    try {
      const baseExpression = new ExcelExpression(cleanFormula);
      this.parseExpressions(baseExpression, 0, cleanFormula, separator);

      const beautified = this.prettyPrint(baseExpression, 0);

      return {
        beautified: (hasEquals ? "=" : "") + beautified,
      };
    } catch (error) {
      // Fallback to original formula if parsing fails
      console.error(error)
      return {
        beautified: formula,
      };
    }
  }
}
