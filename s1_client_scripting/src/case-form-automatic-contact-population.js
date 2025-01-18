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
   * Populates the contact field on the case form when the customer field
   * changes.
   *
   * @param {Object} executionContext   The execution context provided by the
   *                                    form event.
   * @returns {Promise<void>}           A promise that resolves when the contact
   *                                    field is populated
   */
  async function populateContactOnCustomerChange(executionContext) {
    _guardThatExecutionContextIsNotNull(executionContext);

    const formContext = executionContext.getFormContext();
    _guardThatFormIsAssociatedWithTheCaseEntity(formContext);

    const contact = await _getContactLookupValueFromCustomerField(formContext);

    _setCaseContactField(formContext, contact);
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
    _guardThatContactFieldIsNotNull(contactField);
    contactField.setValue(contactLookup);
    contactField.fireOnChange();
  }

  // Throws an error if execution context is null
  function _guardThatExecutionContextIsNotNull(executionContext) {
    if (executionContext === null) {
      throw new Error(
        "The execution context must be passed as the first parameter"
      );
    }
  }

  // Throws an error if form context is null or associated with an entity other
  // than case/incident
  function _guardThatFormIsAssociatedWithTheCaseEntity(formContext) {
    if (
      !formContext ||
      formContext.contextToken?.entityTypeName !== logicalNames.tables.case
    ) {
      throw new Error(
        "Invalid form: This handler should only be used on a case table form"
      );
    }
  }

  // Throws an error if the contact field cannot be found on the form
  function _guardThatContactFieldIsNotNull(contactField) {
    if (contactField === null) {
      throw new Error("The contact field must be present in the form");
    }
  }

  //Return the API
  return {
    populateContactOnCustomerChange,
  };
})();
