const path = require('path');
const WordExtractor = require('../lib/word');

describe('Word file test02.doc', () => {

  const extractor = new WordExtractor();

  it('should match the expected body', (done) => {
    const extract = extractor.extract(path.resolve(__dirname, "data/test02.doc"));
    return extract
      .then((result) => {
        const body = result.getBody();
        expect(body).toMatch(/My name is Ryan/);
        expect(body).toMatch(/create several FKPs for testing purposes/);
        expect(body).toMatch(/dsadasdasdasd/);
        return done();
      });
  });

  it('should return an object of four parts', (done) => {
    const extract = extractor.extractWithStyles(path.resolve(__dirname, "data/test02.doc"));
    return extract.then(result => {
      expect(result['WordDocument']).toBeDefined();
      expect(result['Data']).toBeDefined();
      // expect(result['0Table']).toBeDefined();
      // expect(result['1Table']).toBeDefined();
      done();
    });
  });
});
