"use strict";

this.cr4fd = this.window || {};
this.cr4fd.caseFormAutomaticContactPopulation = (function () {
  // eslint-disable-next-line no-undef
  const _xrm = Xrm;

  //Dictionary containing logical names for entities and entity fields
  const _logicalNames = {
    tables: {
      account: "account",
      contact: "contact",
      case: "incident",
    },
    caseFields: {
      contact: "primarycontactid",
      customer: "customerid",
    },
    accountFields: {
      primaryContact: "primarycontactid",
    },
    contactFields: {
      id: "contactid",
      fullname: "fullname",
    },
  };

  /**
   * Populate the contact field in a case form using the value from the customer
   * field.
   *
   * @param {Object} executionContext  Execution context from a form event
   */
  async function populateContactOnCustomerChange(executionContext) {
    try {
      _guardExecutionContextPassed(executionContext);

      const formContext = _tryReadValidFormContextOrThrow(formContext);

      const contact = await _getContactLookupValueFromCustomerField(
        formContext
      );

      _setCaseContactField(formContext, contact);
    } catch (error) {
      console.error(error);
      _notifyUserOfError(error);
    }
  }

  /**
   * Retrieve the lookup value for a contact based on the customer field.
   * If the customer field references an account entity, it extracts the primary
   * contact where one exists. If the customer field is null or contains a
   * contact, the value of the customer field is simply returned.
   *
   * @param {Object} formContext  The form context object.
   * @returns {Promise<Object>}   A promise that resolves to the contact
   *                              lookup value.
   */
  async function _getContactLookupValueFromCustomerField(formContext) {
    const customerFieldValue = formContext
      .getAttribute(_logicalNames.caseFields.customer)
      .getValue();

    if (
      customerFieldValue &&
      customerFieldValue.length > 0 &&
      customerFieldValue[0].entityType === _logicalNames.tables.account
    ) {
      const accountId = customerFieldValue[0].id;
      return await _getPrimaryContactLookupValueFromAccount(accountId);
    }
    return customerFieldValue;
  }

  /**
   * Retrieves the primary contact lookup value from an account record. Returns
   * null if the primary contact field is not populated
   *
   * @param {string} accountId  The ID of the account to retrieve the primary
   *                            contact from.
   * @returns {Promise<Object|null>}  A promise that resolves to the primary
   *                                  contact lookup value.
   * @throws {Error}  If there is an error retrieving the account record.
   */
  async function _getPrimaryContactLookupValueFromAccount(accountId) {
    try {
      const account = await _xrm.WebApi.retrieveRecord(
        _logicalNames.tables.account,
        accountId,
        _buildSelectsQueryStringForPrimaryContact()
      );
      return _buildPrimaryContactLookupFromAccountRecord(account);
    } catch (error) {
      console.error(error);
      throw new Error(
        "The form may not behave as expected Please reload the form. If the " +
          "problem persists contact an administrator"
      );
    }
  }

  /**
   * Builds a query string to fetch and expand the primary contact from an
   * account record. The query string includes the primary contact and selects
   * the full name field of the contact.
   *
   * @returns {string} The query string to fetch and expand the primary contact.
   */
  function _buildSelectsQueryStringForPrimaryContact() {
    return (
      "?$expand=" +
      _logicalNames.accountFields.primaryContact +
      "($select=" +
      _logicalNames.contactFields.fullname +
      ")"
    );
  }

  /**
   * Format the expanded primary contact field of an account record as a contact
   * lookup value. Returns undefined if the account record or primary contact
   * field are null.
   *
   * @param {Object} accountRecord  The account record containing the primary
   *                                contact field.
   * @returns {Object|null}  A contact lookup object with id, name, and
   *                         entityType properties, or null if no primary
   *                         contact is found.
   */
  function _buildPrimaryContactLookupFromAccountRecord(accountRecord) {
    const contact =
      accountRecord &&
      accountRecord[_logicalNames.accountFields.primaryContact];

    if (contact) {
      return [
        {
          id: contact[_logicalNames.contactFields.id],
          name: contact[_logicalNames.contactFields.fullname],
          entityType: _logicalNames.tables.contact,
        },
      ];
    }
  }

  /**
   * Sets the contact field on the case form to the provided contact lookup
   * value.
   *
   * @param {Object} formContext  The form context object.
   * @param {Object} contactLookup  The lookup value for the contact to be set.
   */
  function _setCaseContactField(formContext, contactLookup) {
    const contactField = formContext.getAttribute(
      _logicalNames.caseFields.contact
    );
    contactField.setValue(contactLookup);
    contactField.fireOnChange();
  }

  /**
   * Validates that the execution context is defined and contains a
   * getFormContext method.
   *
   * @param {Object} executionContext  The execution context to validate.
   * @throws {Error}  Throws an error if the execution context is invalid or
   *                  does not contain a getFormContext method.
   */
  function _guardExecutionContextPassed(executionContext) {
    if (typeof executionContext?.getFormContext !== "function") {
      throw new Error(
        "Invalid execution context. Ensure that execution context is passed " +
          "as the first parameter"
      );
    }
  }

  /**
   * Reads and validates the form context from the execution context.
   *
   * @param {Object} executionContext  The execution context
   * @throws {Error}  Throws an error if the form is not associated with the
   *                  case entity or is missing the contact control
   */
  function _tryReadValidFormContextOrThrow(executionContext) {
    const formContext = executionContext.getFormContext();
    const errorHandler = (message) => {
      throw new Error(`Invalid form configuration: ${message}`);
    };
    _guardFormIsAssociatedWithTheCaseEntity(formContext, errorHandler);
    _guardContactControlIsPresent(formContext, errorHandler);
  }

  /**
   * Validates that the form is associated with the case entity. Calls the error
   * handler if the form is not associated with the contact entity
   *
   * @param {Object} formContext  The form context
   * @param {Function} errorHandler  The function to call with an error message
   *                                 if validation fails
   */
  function _guardFormIsAssociatedWithTheCaseEntity(formContext, errorHandler) {
    if (
      formContext?.contextToken?.entityTypeName !== _logicalNames.tables.case
    ) {
      errorHandler(
        `Form must be associated with ${_logicalNames.tables.case} entity`
      );
    }
  }

  /**
   * Validates that the form contains the contact field control. Calls the error
   * handler if this control is missing
   *
   * @param {Object} formContext  The form context
   * @param {Function} errorHandler  The function to call with an error message
   *                                 if validation fails
   */
  function _guardContactControlIsPresent(formContext, errorHandler) {
    const contactField = formContext.getAttribute(
      _logicalNames.caseFields.contact
    );
    if (!contactField) {
      errorHandler("The contact field control must be present in the form");
    }
  }

  /**
   * Displays an error dialog to the user with a specified error message and
   * details.
   *
   * @param {Error} error  The error object to display
   */
  function _notifyUserOfError(error) {
    const plugInName = populateContactOnCustomerChange.name;
    _xrm.Navigation.openErrorDialog({
      message: `${plugInName} has encountered an error. ${error.message}`,
      details: error.stack,
    });
  }

  // Return the API
  return {
    populateContactOnCustomerChange,
  };
})();
