import categories from '../fixtures/categories.json';

describe('Dashboard Functionality', () => {
  before(() => {
    cy.visit('https://srfclm-test.smartcontract.co.in/#/Login', { timeout: 80000 });
    cy.contains('mat-label', 'Email Id').parents('mat-form-field').find('input').type('tchugh@srtekbox.com');
    cy.contains('mat-label', 'Password').parents('mat-form-field').find('input').type('P@ssw0rd');
    cy.contains('button', 'Log In').click();
    cy.url().should('include', 'https://srfclm-test.smartcontract.co.in/#/Actionables');
  });

  after(() => {
    cy.log('All tests completed');
  });
  Object.entries(categories).forEach(([contractType, fields]) => {

    it(`Validate ${contractType}`, () => {
      cy.get('body').then(($body) => {
      cy.contains('New Contract Request', { timeout: 20000 })
        .should('be.visible')
        .click();

      cy.get('mat-select', { timeout: 20000 })
        .should('exist')
        .first()
        .click();
        cy.contains('mat-option', contractType)
          .should('be.visible')
          .click();

 // Step 4: Validate fields
      cy.wrap(fields).each((field) => {

        cy.log(`Checking field: ${field.name}`);

       cy.contains(field.name, { timeout: 15000, matchCase: false })
          .should('be.visible');

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
});