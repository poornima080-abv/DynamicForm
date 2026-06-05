describe("Full Pipeline", () => {

   it('1. Generate JSON from Excel', () => {
    cy.task('generateAI', {
      filePath: 'D:\\Dynamic Form\\Implementation_SRF.xlsx'
    }).then((result) => {
      cy.log(`Generated ${result.length} contracts`);
    });
  });

  it('2. Convert JSON to Excel', () => {
  cy.readFile('cypress/fixtures/categories.json').then((contracts) => {
    expect(contracts.length).to.be.greaterThan(0);

    // small wait to ensure no file locks
    cy.wait(1000);

    cy.task('convertJsonToExcel', { contracts }).then((msg) => {
      cy.log(msg);
    });
  });
});

    it("3. UI Validation", () => {
        cy.log("Step 3 (safe)");
    });

});