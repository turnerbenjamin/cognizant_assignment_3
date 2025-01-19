//Set publisher namespace
this.cr4fd = this.window || {};

// Initialise namespace for current plugin
this.cr4fd.caseFormAvailableCommunicationChannelsUpdate = (function () {
  // eslint-disable-next-line no-undef
  const xrm = Xrm;

  //Dictionary containing logical names for entities and entity fields
  const logicalNames = {
    tables: {
      case: "incident",
      contact: "contact",
    },
    caseFields: {
      contact: "primarycontactid",
      emailAddress: "emailaddress",
    },
    contactFields: {
      mobilePhoneNumber: "mobilephone",
      emailAddress: "emailaddress1",
      doNotPhone: "donotphone",
      doNotEmail: "donotemail",
    },
    controls: {
      contactAvailableMethodsQuickView:
        "contact_available_contact_methods_view",
    },
  };

  //Dictionary of requirement level options for a control
  const fieldRequirementLevels = {
    required: "required",
    optional: "none",
  };

  /**
   * Updates the available channels section on the form based on the contact's
   * available contact methods.
   *
   * @param {Object} executionContext   The execution context provided by the
   *                                    form event.
   * @returns {Promise<void>}           A promise that resolves when the
   *                                    available channels section is updated.
   */
  async function updateAvailableChannelsSection(executionContext) {
    try {
      _guardExecutionContextValid(executionContext);
      const formContext = executionContext.getFormContext();
      _guardFormContextIsValid(formContext);

      const contact = await _getContactRecordFromContactField(formContext);
      const availableContactMethods = _getContactMethodAvailability(contact);

      _updateContactQuickView(formContext, availableContactMethods);
      _updateCaseEmailField(formContext, availableContactMethods);
    } catch (error) {
      console.error(error);
      _notifyUserOfError(error);
    }
  }

  // Fetch the contact record from the contact lookup field. Returns null if the
  // contact field is not populated
  async function _getContactRecordFromContactField(formContext) {
    const contactFieldValue = formContext
      .getAttribute(logicalNames.caseFields.contact)
      .getValue();
    if (!contactFieldValue || contactFieldValue.length === 0) {
      return null;
    }

    const contactId = contactFieldValue[0].id;

    return await _tryRetrieveContactRecordById(contactId);
  }

  async function _tryRetrieveContactRecordById(contactId) {
    try {
      return await xrm.WebApi.retrieveRecord(
        logicalNames.tables.contact,
        contactId,
        _getSelectsQueryStringForContact()
      );
    } catch (error) {
      console.error(error);
      throw new Error(
        "The form may not behave as expected Please reload the form. If the " +
          "problem persists contact an administrator"
      );
    }
  }

  // Updates the contact quick view. If there are no available communication
  // channeld the view is hidden. Else, a field will be shown within the view
  // for each available contact channel.
  function _updateContactQuickView(formContext, availableContactMethods) {
    const contactControls = _getContactViewControls(formContext);

    contactControls.quickViewControl.setVisible(
      availableContactMethods.mobilePhone ||
        availableContactMethods.emailAddress
    );

    // Note, the controls will be null if no values are present
    contactControls.mobilePhoneControl?.setVisible(
      availableContactMethods.mobilePhone
    );
    contactControls.emailControl?.setVisible(
      availableContactMethods.emailAddress
    );
  }

  // Updates the email field on the case form. This will be shown and made
  // mandatory where no communication channels can be found for a contact. Else,
  // it is hidden and optional
  function _updateCaseEmailField(formContext, availableContactMethods) {
    const isNoContactCommunicationChannelsAvailable =
      !availableContactMethods?.mobilePhone &&
      !availableContactMethods?.emailAddress;

    const requiredLevel = isNoContactCommunicationChannelsAvailable
      ? fieldRequirementLevels.required
      : fieldRequirementLevels.optional;

    formContext
      .getControl(logicalNames.caseFields.emailAddress)
      .setVisible(isNoContactCommunicationChannelsAvailable);
    formContext
      .getAttribute(logicalNames.caseFields.emailAddress)
      .setRequiredLevel(requiredLevel);
  }

  // Returns an object with references to the controls in the contact quick view
  function _getContactViewControls(formContext) {
    const quickViewControl = _tryReadContactQuickViewForm(formContext);

    return {
      quickViewControl,
      mobilePhoneControl: quickViewControl.getControl(
        logicalNames.contactFields.mobilePhoneNumber
      ),
      emailControl: quickViewControl.getControl(
        logicalNames.contactFields.emailAddress
      ),
    };
  }

  function _tryReadContactQuickViewForm(formContext) {
    const quickViewControl = formContext?.ui?.quickForms.get(
      logicalNames.controls.contactAvailableMethodsQuickView
    );
    if (!quickViewControl) {
      throw new Error(
        "Invalid form configuration: Form must contain a contact quick view " +
          "form with the name " +
          `"${logicalNames.controls.contactAvailableMethodsQuickView}"`
      );
    }
    return quickViewControl;
  }

  // Returns an object with keys for each channel and bool values indicating
  // whether the channel is available.
  function _getContactMethodAvailability(contact) {
    const availability = {
      mobilePhone: false,
      emailAddress: false,
    };

    if (contact) {
      availability.mobilePhone = _isContactMethodAvailable(
        contact[logicalNames.contactFields.mobilePhoneNumber],
        contact[logicalNames.contactFields.doNotPhone]
      );
      availability.emailAddress = _isContactMethodAvailable(
        contact[logicalNames.contactFields.emailAddress],
        contact[logicalNames.contactFields.doNotEmail]
      );
    }

    return availability;
  }

  // Returns a bool indicating whether a channel is available. Checks that the
  // channel has a value and that the related preferences allow contact through
  // this channel
  function _isContactMethodAvailable(channelValue, doNotContactPreference) {
    if (
      !channelValue ||
      typeof channelValue !== "string" ||
      typeof doNotContactPreference !== "boolean"
    ) {
      return false;
    }

    return !doNotContactPreference && channelValue.trim() !== "";
  }

  // Build a selects query string for a contact record to return details needed
  // to identify which communication channels are available
  function _getSelectsQueryStringForContact(isNested) {
    var prefix = isNested ? "" : "?";
    var fields = [
      logicalNames.contactFields.emailAddress,
      logicalNames.contactFields.doNotEmail,
      logicalNames.contactFields.mobilePhoneNumber,
      logicalNames.contactFields.doNotPhone,
    ];
    return prefix + "$select=" + fields.join(",");
  }

  function _guardExecutionContextValid(executionContext) {
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

  function _guardFormContextIsValid(formContext) {
    _guardFormIsAssociatedWithCaseEntity(formContext);
    _guardFieldIsPresentInForm(formContext, logicalNames.caseFields.contact);
    _guardFieldIsPresentInForm(
      formContext,
      logicalNames.caseFields.emailAddress
    );
  }

  // Calls the error handler with error detail if the form is not associated
  // with the case form
  function _guardFormIsAssociatedWithCaseEntity(formContext) {
    if (
      formContext?.contextToken?.entityTypeName !== logicalNames.tables.case
    ) {
      throw new Error(
        "Invalid form configuration: Form must be associated with" +
          `${logicalNames.tables.case} entity`
      );
    }
  }

  // Calls the error handler with error detail if the form does not contain the
  // contact field
  function _guardFieldIsPresentInForm(formContext, fieldLogicalName) {
    const necessaryField = formContext.getAttribute(fieldLogicalName);
    if (!necessaryField) {
      throw new Error(
        "Invalid form configuration: Form must contain " +
          `${fieldLogicalName} field`
      );
    }
  }

  // Displays an error to the user with a message
  function _notifyUserOfError(error) {
    const plugInName = updateAvailableChannelsSection.name;
    xrm.Navigation.openErrorDialog({
      message: `${plugInName} has encountered an error. ${error.message}`,
      details: error.stack,
    });
  }

  //Return the API
  return {
    updateAvailableChannelsSection,
  };
})();
