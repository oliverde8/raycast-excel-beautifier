export interface FormulaDoc {
  name: string;
  category: string;
  description: string;
  syntax: string;
  example: string;
}

export const FORMULAS_DOCUMENTATION: Record<string, FormulaDoc> = {
  SUM: {
    name: "SUM",
    category: "Basic",
    description: "Adds values together",
    syntax: "SUM(number1, [number2], ...)",
    example: "=SUM(A1:A10)",
  },
  AVERAGE: {
    name: "AVERAGE",
    category: "Basic",
    description: "Calculates the average of values",
    syntax: "AVERAGE(number1, [number2], ...)",
    example: "=AVERAGE(B1:B10)",
  },
  COUNT: {
    name: "COUNT",
    category: "Basic",
    description: "Counts the number of cells with numbers",
    syntax: "COUNT(value1, [value2], ...)",
    example: "=COUNT(A1:A100)",
  },
  COUNTA: {
    name: "COUNTA",
    category: "Basic",
    description: "Counts non-empty cells",
    syntax: "COUNTA(value1, [value2], ...)",
    example: "=COUNTA(A1:A100)",
  },
  MAX: {
    name: "MAX",
    category: "Basic",
    description: "Returns the maximum value",
    syntax: "MAX(number1, [number2], ...)",
    example: "=MAX(A1:A10)",
  },
  MIN: {
    name: "MIN",
    category: "Basic",
    description: "Returns the minimum value",
    syntax: "MIN(number1, [number2], ...)",
    example: "=MIN(A1:A10)",
  },
  IF: {
    name: "IF",
    category: "Conditional",
    description: "Returns one value if a condition is true, another if false",
    syntax: "IF(logical_test, [value_if_true], [value_if_false])",
    example: "=IF(A1>100, \"High\", \"Low\")",
  },
  IFERROR: {
    name: "IFERROR",
    category: "Conditional",
    description: "Returns a specified value if a formula results in an error",
    syntax: "IFERROR(value, value_if_error)",
    example: "=IFERROR(A1/B1, \"N/A\")",
  },
  IFS: {
    name: "IFS",
    category: "Conditional",
    description: "Checks multiple conditions and returns the corresponding value",
    syntax: "IFS(condition1, value1, [condition2, value2], ...)",
    example: "=IFS(A1>100, \"High\", A1>50, \"Medium\", \"Low\")",
  },
  VLOOKUP: {
    name: "VLOOKUP",
    category: "Lookup",
    description: "Searches for a value in the first column of a range and returns a value in the same row",
    syntax: "VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])",
    example: "=VLOOKUP(E1, Sheet1!A:B, 2, FALSE)",
  },
  INDEX: {
    name: "INDEX",
    category: "Lookup",
    description: "Returns a value from a range at a specified position",
    syntax: "INDEX(array, row_num, [column_num])",
    example: "=INDEX(A1:B10, 3, 2)",
  },
  MATCH: {
    name: "MATCH",
    category: "Lookup",
    description: "Returns the position of an item in a range",
    syntax: "MATCH(lookup_value, lookup_array, [match_type])",
    example: "=MATCH(\"Apples\", A1:A10, 0)",
  },
  CONCATENATE: {
    name: "CONCATENATE",
    category: "Text",
    description: "Joins multiple text strings together",
    syntax: "CONCATENATE(text1, [text2], ...)",
    example: "=CONCATENATE(\"Hello\", \" \", \"World\")",
  },
  UPPER: {
    name: "UPPER",
    category: "Text",
    description: "Converts text to uppercase",
    syntax: "UPPER(text)",
    example: "=UPPER(\"hello\")",
  },
  LOWER: {
    name: "LOWER",
    category: "Text",
    description: "Converts text to lowercase",
    syntax: "LOWER(text)",
    example: "=LOWER(\"HELLO\")",
  },
  TRIM: {
    name: "TRIM",
    category: "Text",
    description: "Removes leading and trailing spaces",
    syntax: "TRIM(text)",
    example: "=TRIM(\" hello \")",
  },
  LEN: {
    name: "LEN",
    category: "Text",
    description: "Returns the length of a text string",
    syntax: "LEN(text)",
    example: "=LEN(\"hello\")",
  },
  ROUND: {
    name: "ROUND",
    category: "Math",
    description: "Rounds a number to a specified number of digits",
    syntax: "ROUND(number, num_digits)",
    example: "=ROUND(3.14159, 2)",
  },
  ABS: {
    name: "ABS",
    category: "Math",
    description: "Returns the absolute value of a number",
    syntax: "ABS(number)",
    example: "=ABS(-10)",
  },
  SQRT: {
    name: "SQRT",
    category: "Math",
    description: "Returns the square root of a number",
    syntax: "SQRT(number)",
    example: "=SQRT(16)",
  },
  TODAY: {
    name: "TODAY",
    category: "Date",
    description: "Returns the current date",
    syntax: "TODAY()",
    example: "=TODAY()",
  },
  NOW: {
    name: "NOW",
    category: "Date",
    description: "Returns the current date and time",
    syntax: "NOW()",
    example: "=NOW()",
  },
  DATE: {
    name: "DATE",
    category: "Date",
    description: "Creates a date from year, month, and day",
    syntax: "DATE(year, month, day)",
    example: "=DATE(2024, 11, 13)",
  },
};

export function getFormulaDoc(formulaName: string): FormulaDoc | undefined {
  return FORMULAS_DOCUMENTATION[formulaName.toUpperCase()];
}

export function getAllFormulasInCategory(category: string): FormulaDoc[] {
  return Object.values(FORMULAS_DOCUMENTATION).filter((doc) => doc.category === category);
}

export function getAllCategories(): string[] {
  const categories = new Set(Object.values(FORMULAS_DOCUMENTATION).map((doc) => doc.category));
  return Array.from(categories).sort();
}
