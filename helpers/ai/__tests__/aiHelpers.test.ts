import { analyzeData, naturalLanguageSearch, generateRecommendations, generateSmartReport } from '../aiHelpers';

describe('AI helpers offline-by-default', () => {
  const headers = ['назва','кількість'];
  const data = [ ['a',1], ['b',2] ];

  it('analyzeData returns offline notice when AI disabled', async () => {
    delete (process.env as any).OPENAI_API_KEY;
    const res = await analyzeData(data, headers);
    expect(res).toMatch(/AI вимкнено|офлайн/i);
  });

  it('naturalLanguageSearch returns offline explanation when AI disabled', async () => {
    delete (process.env as any).OPENAI_API_KEY;
    const res = await naturalLanguageSearch('a', data, headers);
    expect(res.explanation).toMatch(/офлайн/i);
  });

  it('generateRecommendations returns offline-mode message', async () => {
    delete (process.env as any).OPENAI_API_KEY;
    const res = await generateRecommendations(data, headers);
    expect(res[0]).toMatch(/офлайн/i);
  });

  it('generateSmartReport returns basic report when AI disabled', async () => {
    delete (process.env as any).OPENAI_API_KEY;
    const res = await generateSmartReport(data, headers);
    expect(res.analysis).toMatch(/офлайн/i);
  });
});

