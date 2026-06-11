describe("Full Pipeline", () => {

  it('1. Excel → JSON (no AI)', () => {
    cy.task('generateAI', {
      filePath: 'D:\\Dynamic Form\\Implementation_SRF.xlsx'
    }).then((contracts) => {
      cy.log(`Generated ${contracts.length} contracts`);
    });
  });

  it('2. JSON → Excel (AI visibility)', () => {
    cy.readFile('cypress/fixtures/categoriesAI.json').then((contracts) => {
      cy.task('convertJsonToExcelAI', { contracts }, { timeout: 3600000 }).then((msg) => {
        cy.log(msg);
      });
    });
  });

  it("3. UI Validation", () => {
    cy.log("Step 3 (safe)");
  });

});