
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

  const contractCapture = {};
  function shouldShowField(field, context = "request") {

    const visibility = String(field.Visibility || "").trim().toLowerCase();
    const rule = String(field["Business Rule"] || "").toLowerCase();

    if (visibility === "no") {
      return false;
    }

    if (context === "request") {

      if (rule.includes("don't show this field on request form")) {
        return false;
      }

      if (rule.includes("do not show")) {
        return false;
      }
    }

    if (context === "publish") {
      if (rule.includes("mandatory")) {
        return true;
      }
    }

    return visibility === "yes";
  }

  function waitForLoader() {
    cy.get('body').then(($body) => {
      if ($body.find('#spinner').length > 0) {
        cy.get('#spinner', { timeout: 15000 })
          .should('not.exist');
      }
    });
  }
  function isFixedUploadField(field) {
    const mandatory = String(field["Mandatory"] || "").toLowerCase();
    const rule = String(field["Business Rule"] || "").toLowerCase();
    return mandatory.includes("yes") || rule.includes("mandatory when");
  }
  it('Validate all contract types and field visibility', () => {

    cy.fixture('categories.json').then((meta) => {
      meta.forEach((contract) => {

        const contractType = contract['Contract Type'];
        waitForLoader();
        //  cy.log(`Testing Contract Type: ${contractType}`);

        cy.get('body').click(0, 0);
        cy.get('body').type('{esc}');
        cy.get('mat-select[title="Contract Type"]').filter(':visible').first().click();
        cy.get('.cdk-overlay-pane', { timeout: 10000 }).should('exist');
        cy.contains('mat-option', contractType, { timeout: 10000 }).click({ force: true });

        cy.wait(1000);

        const expectedFields = contract.fields.filter(field =>
          shouldShowField(field, "request")
        );


        expectedFields.forEach((field) => {

          const fieldName = field["Field Name"];
          const fieldType = field["Field Input Type"];
          const fieldValue = String(field["Value"] || "").trim();

          cy.log(`Checking field: ${fieldName}`);
          const firstOption = (val) => val.split('|')[0].trim();

          const contractType = String(contract['Contract Type'] || "").trim();
          if (!contractCapture[contractType]) {
            contractCapture[contractType] = [];
          }
          contractCapture[contractType].push({
            fieldName: fieldName,
            type: fieldType,
            value: fieldValue
          });

          if (fieldType === "Upload") {
            if (isFixedUploadField(field)) {
              cy.contains(fieldName).should('exist');
              cy.get('input[type="file"]')
                .last().selectFile('cypress/fixtures/File/adoc1.pdf', { force: true });
            } else {

              cy.get('mat-select[title="Enter Document Name"]')
                .first()
                .click({ force: true });

              cy.contains('mat-option', fieldName)
                .should('exist')
                .click({ force: true });

              cy.get('input[type="file"]')
                .last()
                .selectFile('cypress/fixtures/File/adoc1.pdf', { force: true });
            }
          } else if (fieldType === "Selection" || fieldType === "Dropdown" || fieldType === "Select") {

            cy.contains('mat-label', fieldName).should('exist').then(() => {
              if (fieldValue && !fieldValue.includes('Matrix') && fieldValue.trim() !== '') {

                // split by | and take first clean option
                const optionToSelect = fieldValue.split('|')[0].trim();
                cy.contains('mat-label', fieldName)
                  .closest('mat-form-field, div')
                  .find('mat-select')
                  .click({ force: true });

                cy.get('.cdk-overlay-pane mat-option', { timeout: 10000 })
                  .contains(optionToSelect)
                  .click({ force: true });
              }
            });

          } else if (fieldType === "Text" || fieldType === "Number" || fieldType === "Email") {

            cy.contains('mat-label', fieldName).should('exist').then(() => {
              // default test values by type
              const fillValue = fieldValue && fieldValue.trim() !== '' && !fieldValue.includes('|')
                ? fieldValue
                : fieldType === 'Email' ? 'test@test.com'
                  : fieldType === 'Number' ? '1234567890'
                    : 'Test Value';

              cy.get(`input[title="${fieldName}"], textarea[title="${fieldName}"]`)
                .first()
                .then(($el) => {
                  if ($el.length && !$el.is(':disabled')) {
                    cy.wrap($el)
                      .clear({ force: true })
                      .type(fillValue, { force: true });
                  }
                });
            });

          } else if (fieldType === "Multi Lines of Text" || fieldType === "Textarea") {

            cy.contains('mat-label', fieldName).should('exist').then(($label) => {
              cy.wrap($label)
                .closest('div')
                .find('textarea')
                .then(($el) => {
                  if ($el.length && !$el.is(':disabled')) {
                    cy.wrap($el)
                      .clear({ force: true })
                      .type('Test input', { force: true });
                  }
                });
            });

          }
          /*  else if (fieldType === "Date" || fieldType === "Select Date / Auto Populate") {
              cy.get('body').click(0, 0);
              cy.get('body').type('{esc}');
              cy.contains('mat-label', fieldName).should('exist').then(($el) => {
                cy.get(`input[title="${fieldName}"]`).then(($el) => {
                  if ($el.length && !$el.is(':disabled')) {
                    cy.wrap($el).closest('.mat-mdc-form-field-infix').siblings('.mat-mdc-form-field-icon-suffix').find('button').click({ force: true });
                    cy.get('.mat-calendar-body-cell').contains(15).click({ force: true });
                     cy.get('body').click(0, 0);
                     cy.get('body').type('{esc}');
                    //cy.get('.cdk-overlay-backdrop').should('not.exist');
                  }
                });
              });
  
            } 
            */
          else if (fieldType === "Fixed") {

            cy.contains(fieldName).then(($label) => {
              cy.wrap($label).parent().then(($field) => {
                // Dropdown
                if ($field.find('mat-select').length > 0) {

                  cy.wrap($field)
                    .find('mat-select')
                    .click({ force: true });

                  cy.get('mat-option')
                    .not('.mat-mdc-option-disabled')
                    .first()
                    .click({ force: true });
                }

                // Textbox
                else if ($field.find('input, textarea').length > 0) {

                  cy.wrap($field)
                    .find('input, textarea')
                    .first()
                    .then(($input) => {

                      if (!$input.is(':disabled')) {
                        cy.wrap($input)
                          .clear({ force: true })
                          .type('Test Data', { force: true });
                      }
                    });
                }

              });
            }
            )
          } else {

            cy.contains('mat-label', fieldName).should('exist');

          }

        });
      });
      cy.then(() => {
        cy.writeFile(
          'cypress/fixtures/contract-output.json',
          contractCapture);

      });
    });
  });
});