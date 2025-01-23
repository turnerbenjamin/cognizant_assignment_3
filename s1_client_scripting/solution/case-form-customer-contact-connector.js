"use strict";

this.cr4fd = this.window || {};
this.cr4fd.caseFormCustomerContactConnector = (function () {
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

  //Dictionary of requirement level options for a control
  const _requirementLevels = {
    required: "required",
    optional: "none",
  };

  /**
   * If the customer field contains an account, and that account has a primary
   * contact. This handler will populate the contact field with the primary
   * account. If the customer is a contact or null the contact field will be set
   * to null
   *
   * @param {Object} executionContext  Execution context passed as a first
   *                                   parameter for a form event
   */
  async function populateContactOnCustomerChange(executionContext) {
    try {
      _guardExecutionContextIsValid(executionContext);
      const formContext = _tryReadValidFormContextOrThrow(executionContext);

      const contact = await _getContactLookupValueFromCustomerField(
        formContext
      );

      _setCaseContactField(formContext, contact);
    } catch (error) {
      console.error(error);
      _notifyUserOfError(error, populateContactOnCustomerChange.name);
    }
  }

  /**
   * Updates the visibility and requirement level of the contact control in a
   * case form based on the value of the customer field.
   *
   * If the customer is a contact the field is hidden, else it is visible
   * If the customer is an account the field is required, else it is optional
   *
   * @param {Object} executionContext  Execution context passed as a first
   *                                   parameter for a form event
   */
  async function updateContactField(executionContext) {
    try {
      _guardExecutionContextIsValid(executionContext);
      const formContext = _tryReadValidFormContextOrThrow(executionContext);

      _updateContactFieldControl(formContext);
    } catch (error) {
      console.error(error);
      _notifyUserOfError(error, updateContactField.name);
    }
  }

  /**
   * Retrieve the lookup value for a contact based on the customer field.
   * If the customer field references an account entity, it extracts the primary
   * contact where one exists. For all other situations, null is returned
   *
   * @param {Object} formContext  The form context object.
   * @returns {Promise<Object|null>}   A promise that resolves to the contact
   *                                   lookup value.
   */
  async function _getContactLookupValueFromCustomerField(formContext) {
    const customerFieldValue = _readCustomerField(formContext);
    if (customerFieldValue?.entityType !== _logicalNames.tables.account) {
      return null;
    }
    return await _getPrimaryContactLookupValueFromAccount(
      customerFieldValue.id
    );
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
    const account = await _xrm.WebApi.retrieveRecord(
      _logicalNames.tables.account,
      accountId,
      _buildSelectsQueryStringForPrimaryContact()
    );
    return _buildPrimaryContactLookupFromAccountRecord(account);
  }

  /**
   * Builds a query string to fetch and expand the primary contact from an
   * account record.
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
    return null;
  }

  /**
   * Sets the contact field on the case form to the provided contact lookup
   * value and fires an on change event on this field.
   *
   * @param {Object} formContext  The form context object.
   * @param {Object|null} contactLookup  The lookup value for the contact to be set.
   */
  function _setCaseContactField(formContext, contactLookup) {
    const contactField = formContext.getAttribute(
      _logicalNames.caseFields.contact
    );
    contactField.setValue(contactLookup);
    contactField.fireOnChange();
  }

  /**
   * Main control for the logic updating the visibility and requirement
   * level of the contact field
   *
   * If the customer is a contact the field is hidden, else it is visible
   * If the customer is an account the field is required, else it is optional
   *
   * @param {Object} formContext  The form context object.
   */
  function _updateContactFieldControl(formContext) {
    const customerFieldValue = _readCustomerField(formContext);
    const customerType = customerFieldValue?.entityType;

    const isVisible = customerType !== _logicalNames.tables.contact;
    const isRequired = customerType === _logicalNames.tables.account;

    _setContactFieldVisibility(formContext, isVisible);
    _setContactFieldIsRequired(formContext, isRequired);
  }

  /**
   * Sets the visibility of the contact field on the case form.
   *
   * @param {Object} formContext  The form context object
   * @param {boolean} isVisible  A boolean indicating whether the contact field
   *                             should be visible
   */
  function _setContactFieldVisibility(formContext, isVisible) {
    const contactControl = formContext.getControl(
      _logicalNames.caseFields.contact
    );
    contactControl?.setVisible(isVisible);
  }

  /**
   * Sets the requirement level of the contact field on the case form.
   *
   * @param {Object} formContext  A valid form context containing the contact
   *                              control
   * @param {boolean} isRequired  A boolean indicating whether the contact
   *                               field should be required
   */
  function _setContactFieldIsRequired(formContext, isRequired) {
    const requiredLevel = isRequired
      ? _requirementLevels.required
      : _requirementLevels.optional;

    const contactAttribute = formContext.getAttribute(
      _logicalNames.caseFields.contact
    );
    contactAttribute.setRequiredLevel(requiredLevel);
  }

  /**
   * Reads the customer field value from the form context.
   *
   * @param {Object} formContext  The form context object.
   * @returns {Object|null}  The customer field value if it exists, otherwise
   *                         null.
   */
  function _readCustomerField(formContext) {
    const customerFieldValue = formContext
      .getAttribute(_logicalNames.caseFields.customer)
      .getValue();

    if (customerFieldValue && customerFieldValue.length > 0) {
      return customerFieldValue[0];
    }
    return null;
  }

  /**
   * Validates that the execution context is defined and contains a
   * getFormContext method.
   *
   * @param {Object} executionContext  The execution context to validate
   * @throws {Error}  Throws an error if the execution context is invalid or
   *                  does not contain a getFormContext method
   */
  function _guardExecutionContextIsValid(executionContext) {
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
   * @throws {Error}  Throws an error the form is not associated with the case
   *                  entity or is missing the contact control
   */
  function _tryReadValidFormContextOrThrow(executionContext) {
    const formContext = executionContext?.getFormContext();
    const errorHandler = (message) => {
      throw new Error(`Invalid form configuration: ${message}`);
    };

    _guardFormIsAssociatedWithTheCaseEntity(formContext, errorHandler);
    _guardContactControlIsPresent(formContext, errorHandler);

    return formContext;
  }

  /**
   * Validates that the form is associated with the case entity. Calls the error
   * handler if the form is not associated with the case entity
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
    const contactField = formContext?.getControl(
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
  function _notifyUserOfError(error, handlerName) {
    _xrm?.Navigation?.openErrorDialog({
      message: `${handlerName} has encountered an error. ${error.message}`,
      details: error.stack,
    });
  }

  // Return the API
  return {
    populateContactOnCustomerChange,
    updateContactField,
  };
})();
