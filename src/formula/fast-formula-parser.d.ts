declare module 'fast-formula-parser' {
  export default class FormulaParser {
    constructor(config?: {
      onCell?: (ref: { sheet?: string; row: number; col: number }) => unknown;
      onRange?: (ref: { sheet?: string; from: { row: number; col: number }; to: { row: number; col: number } }) => unknown[][];
    });
    parse(formula: string, position: { sheet: string; row: number; col: number }): unknown;
    static FormulaError: abstract new (...args: never[]) => { error: string };
  }
}
