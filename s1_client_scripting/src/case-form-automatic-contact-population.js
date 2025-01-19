this.cr4fd = this.window || {};
this.cr4fd.caseFormAutomaticContactPopulation = (function () {
  // eslint-disable-next-line no-undef
  const xrm = Xrm;

  //Dictionary containing logical names for entities and entity fields
  const logicalNames = {
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
   * Populates the contact field on the case form using the customer field.
   *
   * @param {Object} executionContext   The execution context provided by the
   *                                    form event.
   * @returns {Promise<void>}           A promise that resolves when the contact
   *                                    field is populated
   */
  async function populateContactOnCustomerChange(executionContext) {
    try {
      _validateExecutionContext(executionContext);
      const formContext = executionContext.getFormContext();
      _validateFormContext(formContext);

      const contact = await _getContactLookupValueFromCustomerField(
        formContext
      );

      _setCaseContactField(formContext, contact);
    } catch (error) {
      console.error(error);
      _notifyUserOfError(error);
    }
  }

  // This function is responsible for returning a lookup value for a contact
  // based on the customer field. If the customer field references an account
  // entity, a function is called to extract the primary contact where one
  // exists. If the customer field is null or contains a contact, the value of
  // the customer field is simply returned
  async function _getContactLookupValueFromCustomerField(formContext) {
    const customerFieldValue = formContext
      .getAttribute(logicalNames.caseFields.customer)
      .getValue();

    if (
      customerFieldValue &&
      customerFieldValue.length > 0 &&
      customerFieldValue[0].entityType === logicalNames.tables.account
    ) {
      const accountId = customerFieldValue[0].id;
      return await _getPrimaryContactLookupValueFromAccount(accountId);
    }
    return customerFieldValue;
  }

  // Fetches an account and returns a lookup value for its primary contact. If
  // the account is not found or the primary contact is not populated, null is
  // returned
  async function _getPrimaryContactLookupValueFromAccount(accountId) {
    const account = await xrm.WebApi.retrieveRecord(
      logicalNames.tables.account,
      accountId,
      _getSelectsQueryStringForPrimaryContact()
    );
    return _buildPrimaryContactLookupFromAccountRecord(account);
  }

  // Formats the primary contact field of an account record as a contact lookup
  // value. Returns null if the accountRecord or primary contact field are null
  function _buildPrimaryContactLookupFromAccountRecord(accountRecord) {
    const contact =
      accountRecord && accountRecord[logicalNames.accountFields.primaryContact];

    if (!contact) {
      return null;
    }

    return [
      {
        id: contact[logicalNames.contactFields.id],
        name: contact[logicalNames.contactFields.fullname],
        entityType: logicalNames.tables.contact,
      },
    ];
  }

  //Build a query string to fetch and expand the primary contact from an
  // account record
  function _getSelectsQueryStringForPrimaryContact() {
    return (
      "?$expand=" +
      logicalNames.accountFields.primaryContact +
      "($select=" +
      logicalNames.contactFields.fullname +
      ")"
    );
  }

  //Set the contact field to the contact lookup value provided
  function _setCaseContactField(formContext, contactLookup) {
    const contactField = formContext.getAttribute(
      logicalNames.caseFields.contact
    );
    contactField.setValue(contactLookup);
    contactField.fireOnChange();
  }

  // Validate that the execution context is valid and contains a getFormContext
  // method. Errors thrown are left to the global handler
  function _validateExecutionContext(executionContext) {
    const errorHandler = (message) => {
      throw new Error(`Invalid execution context: ${message}`);
    };
    _guardExecutionContextPresent(executionContext, errorHandler);
    _guardFormContextAccessible(executionContext, errorHandler);
  }

  // Calls the error handler with error detail if execution context is null
  function _guardExecutionContextPresent(executionContext, errorHandler) {
    if (!executionContext) {
      errorHandler(
        "The execution context must be passed as the first parameter"
      );
    }
  }

  // Calls the error handler with error detail if formContext is not accessible
  // from the execution context
  function _guardFormContextAccessible(executionContext, errorHandler) {
    if (typeof executionContext?.getFormContext !== "function") {
      errorHandler(
        "getFormContext is not accessible from the execution context. Ensure " +
          "that execution context is passed as the first parameter"
      );
    }
  }

  // Validates that the form context relates to the case entity and that the
  // contact field is present
  function _validateFormContext(formContext) {
    const errorHandler = (message) => {
      throw new Error(`Invalid form configuration: ${message}`);
    };
    _guardFormIsAssociatedWithTheCaseEntity(formContext, errorHandler);
    _guardContactControlIsPresent(formContext, errorHandler);
  }

  // Calls the error handler with error detail if the form is not associated
  // with the case form
  function _guardFormIsAssociatedWithTheCaseEntity(formContext, errorHandler) {
    if (
      formContext?.contextToken?.entityTypeName !== logicalNames.tables.case
    ) {
      errorHandler(
        `Form must be associated with ${logicalNames.tables.case} entity`
      );
    }
  }

  // Calls the error handler with error detail if the form does not contain the
  // contact field
  function _guardContactControlIsPresent(formContext, errorHandler) {
    const contactField = formContext.getAttribute(
      logicalNames.caseFields.contact
    );
    if (!contactField) {
      errorHandler("The contact field control must be present in the form");
    }
  }

  // Displays an error to the user with a message
  function _notifyUserOfError(error) {
    const plugInName = populateContactOnCustomerChange.name;
    xrm.Navigation.openErrorDialog({
      message: `${plugInName} has encountered an error. ${error.message}`,
      details: error.stack,
    });
  }

  //Return the API
  return {
    populateContactOnCustomerChange,
  };
})();
