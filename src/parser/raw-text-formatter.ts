// Excel Formula Formatter - Pure text formatting logic

import {
  ExcelExpression,
  SubExpression,
  FormulaExpr,
  FormattingOptions,
  DEFAULT_FORMATTING_OPTIONS
} from "./types";

export class ExcelRawTextFormatter {
  private options: FormattingOptions;

  constructor(options: Partial<FormattingOptions> = {}) {
    this.options = { ...DEFAULT_FORMATTING_OPTIONS, ...options };
  }

  format(expression: ExcelExpression, hasEqualsSign: boolean = false): string {
    const result = this.prettyPrint(expression, 0);
    return (hasEqualsSign ? "=" : "") + result;
  }

  private prettyPrint(expression: ExcelExpression, depth: number): string {
    const children = expression.getChilds();

    // If no children, return the original content with operator spacing
    if (children.length === 0) {
      return this.options.useOperatorSpacing 
        ? this.addOperatorSpacing(expression.original || "")
        : expression.original || "";
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
          const content = this.options.useOperatorSpacing 
            ? this.addOperatorSpacing(child.original || "")
            : child.original || "";
          result += content;
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

  private formatFunction(func: FormulaExpr, depth: number): string {
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
          result += this.formatFunction(param, depth + 1);
        } else if (param instanceof SubExpression) {
          result += this.formatSubExpression(param, depth);
        } else {
          const paramContent = param.getChilds().length === 0 
            ? (this.options.useOperatorSpacing ? this.addOperatorSpacing(param.original || "") : param.original || "")
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

      // Add indentation and nesting indicator for parameters
      result += this.indent(depth + 1);
      if (this.options.useNestingIndicators) {
        result += this.getNestingIndicator(depth + 1, isLast);
      }

      if (param instanceof FormulaExpr) {
        // For nested functions, don't add extra indicator since formatFunction will handle it
        const functionResult = this.formatFunction(param, depth + 1);
        result += functionResult;
      } else if (param instanceof SubExpression) {
        const subResult = this.formatSubExpression(param, depth + 1);
        result += subResult;
      } else {
        // For parameter content, preserve original or format children
        const paramContent = param.getChilds().length === 0 
          ? (this.options.useOperatorSpacing ? this.addOperatorSpacing(param.original || "") : param.original || "")
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

  private shouldFormatInline(func: FormulaExpr): boolean {
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
    if (params.length <= (this.options.maxInlineParams || 3)) {
      let totalLength = 0;
      for (const param of params) {
        if (param instanceof FormulaExpr || (param instanceof SubExpression && this.containsFunction(param))) {
          return false;
        }
        const content = param.original || this.prettyPrint(param, 0);
        totalLength += content.length;
        if (totalLength > (this.options.maxInlineLength || 40)) {
          return false;
        }
      }
      return totalLength <= (this.options.maxInlineLength || 40);
    }

    return false;
  }

  private formatSubExpression(subExpr: SubExpression, depth: number): string {
    const content = this.prettyPrint(subExpr, depth);

    // If it's simple, keep on one line
    if (!content.includes('\n') && content.length < 50 && !this.containsFunction(subExpr)) {
      return "(" + content + ")";
    }

    // Complex sub-expression gets multi-line formatting
    return "(\n" + 
      this.indent(depth + 1) + content + "\n" +
      this.indent(depth) + ")";
  }

  private containsFunction(expr: ExcelExpression): boolean {
    if (expr instanceof FormulaExpr) {
      return true;
    }
    return expr.getChilds().some(child => this.containsFunction(child));
  }

  private needsSpacing(current: ExcelExpression, next: ExcelExpression): boolean {
    const currentText = current.original?.trim();
    const nextText = next.original?.trim();

    if (!currentText || !nextText) return false;

    const operators = ['+', '-', '*', '/', '=', '<', '>', '<=', '>=', '<>', '&'];
    return operators.includes(currentText) || operators.includes(nextText);
  }

  private indent(depth: number): string {
    const indentChar = " ".repeat(this.options.indentSize || 4);
    return indentChar.repeat(depth);
  }

  private getNestingIndicator(depth: number, isLast: boolean = false): string {
    if (depth === 0) return "";

    // Create a simple but consistent nesting indicator
    const prefix = "  ".repeat(Math.max(0, depth - 1)); // 2 spaces per level

    if (depth === 1) {
      return isLast ? "└─ " : "├─ ";
    } else {
      return prefix + (isLast ? "└─ " : "├─ ");
    }
  }

  private addOperatorSpacing(text: string): string {
    if (!text) return text;

    // Add spacing around mathematical and comparison operators
    let result = text;

    // Define operators that need spacing
    const operators = [
      { op: '<=', replacement: ' <= ' },
      { op: '>=', replacement: ' >= ' },
      { op: '<>', replacement: ' <> ' },
      { op: '!=', replacement: ' != ' },
      { op: '==', replacement: ' == ' },
      { op: '=', replacement: ' = ' },
      { op: '<', replacement: ' < ' },
      { op: '>', replacement: ' > ' },
      { op: '+', replacement: ' + ' },
      { op: '-', replacement: ' - ' },
      { op: '*', replacement: ' * ' },
      { op: '/', replacement: ' / ' },
      { op: '&', replacement: ' & ' },
    ];

    for (const { op, replacement } of operators) {
      // Use regex to avoid spacing already spaced operators and operators in strings
      const regex = new RegExp(`(?<![\\s<>=!])\\${op.split('').join('\\')}(?![\\s<>=!])`, 'g');
      result = result.replace(regex, replacement);
    }

    // Clean up multiple spaces
    result = result.replace(/\s{2,}/g, ' ');

    // Handle special cases - don't add space around operators in certain contexts
    // Fix negative numbers (don't space before minus if it's at start or after operator/comma/parenthesis)
    result = result.replace(/(^|[+\-*/=<>!&,(])\s*-\s*(\d)/g, '$1-$2');

    // Fix cell ranges (A1:B10 should not have spaces around colon)
    result = result.replace(/([A-Z]+\$?\d+)\s*:\s*([A-Z]+\$?\d+)/g, '$1:$2');

    // Fix absolute references (don't space around $ in cell references)
    result = result.replace(/\$\s*([A-Z]+)\s*\$?\s*(\d+)/g, '$$1$$2');
    result = result.replace(/([A-Z]+)\s*\$\s*(\d+)/g, '$1$$2');

    return result.trim();
  }
}