// Excel Formula Parser and Beautifier - Main Entry Point

import { ExcelFormulaParser } from "./parser";
import { ExcelRawTextFormatter } from "./raw-text-formatter";
import { FormattingOptions } from "./types";

export class ExcelFormulaBeautifier {
  static rawText(formula: string, options?: Partial<FormattingOptions>): string {
    try {
      // Parse the formula
      const hasEquals = formula.trim().startsWith("=");
      const parsedExpression = ExcelFormulaParser.parse(formula);

      // Format the parsed expression
      const formatter = new ExcelRawTextFormatter(options);
      return formatter.format(parsedExpression, hasEquals);
    } catch (error) {
      // Fallback to original formula if parsing fails
      console.error("Formatting error:", error);
      return formula;
    }
  }
}
