// Core types and interfaces for Excel Formula Parser

import { FormulaTypes } from "./formula-types";

export class ExcelExpression {
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

export class SubExpression extends ExcelExpression {}

export class FormulaExpr extends ExcelExpression {
  readonly formula: FormulaTypes;

  constructor(original: string, formula: FormulaTypes) {
    super(original);
    this.formula = formula;
  }
}

export class SimpleExpression extends ExcelExpression {}

export class OperatorExpression extends ExcelExpression {
  readonly operator: string;

  constructor(operator: string) {
    super(operator);
    this.operator = operator;
  }
}

export interface FormattingOptions {
  useNestingIndicators?: boolean;
  useOperatorSpacing?: boolean;
  indentSize?: number;
  maxInlineLength?: number;
  maxInlineParams?: number;
}

export const DEFAULT_FORMATTING_OPTIONS: FormattingOptions = {
  useNestingIndicators: true,
  useOperatorSpacing: true,
  indentSize: 4,
  maxInlineLength: 40,
  maxInlineParams: 3,
};