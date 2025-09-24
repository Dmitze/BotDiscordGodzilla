import { evaluateFormula, validateFormula, setVariable, clearVariables } from '@/utils/formulaProcessor';

describe('FormulaProcessor safe evaluator', () => {
  beforeEach(() => {
    clearVariables();
  });

  it('evaluates basic arithmetic', async () => {
    const res = await evaluateFormula('1+2*3');
    expect(res).toBe(7);
  });

  it('supports allowed functions', async () => {
    const res = await evaluateFormula('round(PI)');
    expect(res).toBe(3);
  });

  it('uses variables from context', async () => {
    setVariable('X', 10);
    const res = await evaluateFormula('X*2');
    expect(res).toBe(20);
  });

  it('rejects disallowed identifiers', async () => {
    const v = validateFormula('process.exit(0)');
    expect(v.isValid).toBe(false);
  });

  it('rejects unknown functions', async () => {
    const v = validateFormula('evil(1)');
    expect(v.isValid).toBe(false);
  });
});


