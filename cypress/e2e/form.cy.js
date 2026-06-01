
describe('Dashboard Functionality', () => {
  before(() => {
    cy.visit('https://srfclm-test.smartcontract.co.in/#/Login', { timeout: 80000 });
    cy.contains('mat-label', 'Email Id').parents('mat-form-field').find('input').type('tchugh@srtekbox.com');
    cy.contains('mat-label', 'Password').parents('mat-form-field').find('input').type('P@ssw0rd');
    cy.contains('button', 'Log In').click();
    cy.url().should('include', 'https://srfclm-test.smartcontract.co.in/#/Actionables');
    cy.contains('Go to Workbench', { timeout: 20000 }).should('be.visible').click();
    cy.contains('New Contract Request', { timeout: 20000 }).should('be.visible').click();

  });

  after(() => {
    cy.log('All tests completed');
  });

  const contractType = "Addendum"; // you can parametrize later

  function loadContractData(type) {
    const filePath = path.join(
      __dirname,
      "../generated",
      `generated_${type}.json`
    );

    return JSON.parse(fs.readFileSync(filePath, "utf8"));
  }

    it("should validate all contract types dynamically", () => {

    // Step 1: load master metadata
    cy.fixture("categories.json").then((meta) => {
  cy.log(JSON.stringify(meta.contractTypes));


      const contractTypes = meta.contractTypes;

      // Step 2: loop ALL contract types dynamically
      contractTypes.forEach((contractType) => {

        cy.log(`Testing contract type: ${contractType}`);


        // open dropdown
        cy.get("mat-select").first().click();

        // select contract type dynamically
        cy.contains("mat-option", contractType, { timeout: 10000 })
          .should("be.visible")
          .click();

        // Step 3: get expected fields for this type
        const expectedFields = meta.fields.filter((field) => {

          return (
            field.visibility?.[contractType] &&
            field.visibility?.[contractType]
              .toString()
              .toLowerCase() === "yes"
          );

        });

        // Step 4: validate UI fields
        expectedFields.forEach((field) => {

          cy.contains(field.name, { timeout: 10000 })
            .should("be.visible");

        });

      });

    });

  });

});