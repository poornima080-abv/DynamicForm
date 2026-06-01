import categories from '../fixtures/categories.json';

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
  Object.entries(categories).forEach(([contractType, fields]) => {
  it(`Validate ${contractType}`, () => {
    cy.get('mat-select[title="Contract Type"]')
  .filter(':visible')
  .first()
  .click();

   cy.get('mat-option')
  .should('have.length.greaterThan', 0);

cy.contains('mat-option', contractType, { timeout: 10000 })
  .click(); // Step 4: Validate fields
    cy.wrap(fields).each((field) => {
      cy.log(`Checking field: ${field.name}`);
      cy.contains(field.name, { timeout: 15000, matchCase: false }).should('be.visible');

      // Validate field type (DOM-based, not formControl-based)
      cy.contains(field.name, { matchCase: false })
        .parents('mat-form-field, div')
        .then(($el) => {

          const type = (field.type || '').toLowerCase();

          if (type === 'text') {
            cy.wrap($el).find('input')
              .should('exist');
          }

          if (type === 'selection') {
            cy.wrap($el).find('mat-select')
              .should('exist');
          }

          if (type === 'textarea') {
            cy.wrap($el).find('textarea')
              .should('exist');
          }

        });

    });
  });
  });

});