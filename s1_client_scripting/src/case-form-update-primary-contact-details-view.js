"use strict";

//Set publisher namespace
this.cr4fd = this.window || {};

// Initialise namespace for current plugin
this.cr4fd.caseFormAvailableCommunicationChannelsUpdate = (function () {
  // eslint-disable-next-line no-undef
  const _xrm = Xrm;

  //Dictionary containing logical names for entities and entity fields
  const _logicalNames = {
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
   * Updates a case form to display available channels of communication based on
   * the contact value. Shows and requires the email field if no available
   * channels can be derived from the contact field
   *
   * @param {Object} executionContext   The execution context provided by the
   *                                    form event.
   */
  async function updateAvailableChannelsSection(executionContext) {
    try {
      _guardExecutionContextPassed(executionContext);
      const formContext = _tryReadValidFormContextOrThrow(executionContext);

      const contact = await _getContactRecordFromContactField(formContext);
      const availableContactMethods = _getContactMethodAvailability(contact);

      _updateContactQuickViewVisiblity(formContext, availableContactMethods);
      _updateCaseEmailField(formContext, availableContactMethods);
    } catch (error) {
      console.error(error);
      _notifyUserOfError(error);
    }
  }

  /**
   * Fetches the contact record from the contact lookup field.
   * Returns undefined if the contact field is not populated.
   *
   * @param {Object} formContext  The form context provided by the execution
   *                              context.
   * @returns {Object|undefined}  The contact record if the contact field is
   *                              populated, otherwise undefined.
   */
  async function _getContactRecordFromContactField(formContext) {
    const contactFieldValue = formContext
      .getAttribute(_logicalNames.caseFields.contact)
      .getValue();

    if (contactFieldValue && contactFieldValue.length > 0) {
      const contactId = contactFieldValue[0].id;
      return await _tryRetrieveContactRecordById(contactId);
    }
  }

  /**
   * Retrieves a contact record by its ID.
   *
   * @param {string} contactId  The ID of the contact record to retrieve.
   * @returns {Promise<Object>}  A promise that resolves to the contact record.
   * @throws {Error}  If the contact record retrieval fails.
   */
  async function _tryRetrieveContactRecordById(contactId) {
    try {
      return await _xrm.WebApi.retrieveRecord(
        _logicalNames.tables.contact,
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

  /**
   * Builds a query string to retrieve fields needed to identify which
   * communication channels are available for a contact.
   *
   * @returns {string} The query string
   */
  function _getSelectsQueryStringForContact() {
    var fields = [
      _logicalNames.contactFields.emailAddress,
      _logicalNames.contactFields.doNotEmail,
      _logicalNames.contactFields.mobilePhoneNumber,
      _logicalNames.contactFields.doNotPhone,
    ];
    return "?$select=" + fields.join(",");
  }

  /**
   * Returns an object indicating the availability of contact methods for a
   * given contact.
   *
   * @param {Object} contact  The contact object containing contact method
   *                          details.
   * @param {string} contact.mobilePhoneNumber
   * @param {boolean} contact.doNotPhone
   * @param {string} contact.emailAddress
   * @param {boolean} contact.doNotEmail
   * @returns {Object}  An object with keys for each channel (mobilePhone,
   *                    emailAddress) and boolean values indicating whether the
   *                    channel is available.
   */
  function _getContactMethodAvailability(contact) {
    const availability = {
      mobilePhone: false,
      emailAddress: false,
    };

    if (contact) {
      availability.mobilePhone = _isContactMethodAvailable(
        contact[_logicalNames.contactFields.mobilePhoneNumber],
        contact[_logicalNames.contactFields.doNotPhone]
      );
      availability.emailAddress = _isContactMethodAvailable(
        contact[_logicalNames.contactFields.emailAddress],
        contact[_logicalNames.contactFields.doNotEmail]
      );
    }

    return availability;
  }

  /**
   * Checks if a contact method is available based on its value and the contact
   * preference.
   *
   * @param {string} channelValue  The value of the contact method, e.g. email
   * @param {boolean} doNotContactPreference  A boolean indicating whether the
   *                                          contact method should not be used
   * @returns {boolean}  Returns true if the contact method is available,
   *                     otherwise false.
   */
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

  /**
   * Updates the visibility of fields within the contact quick view form to
   * show only fields with available contact methods. The quick view form itself
   * will be hidden if no available contact methods are available
   *
   * @param {Object} formContext  The form context object.
   * @param {Object} availableContactMethods  An object containing the
   *                                          available contact methods.
   * @param {boolean} availableContactMethods.mobilePhone
   * @param {boolean} availableContactMethods.emailAddress
   */
  function _updateContactQuickViewVisiblity(
    formContext,
    availableContactMethods
  ) {
    const contactControls = _getContactViewControls(formContext);

    contactControls.quickViewControl.setVisible(
      availableContactMethods.mobilePhone ||
        availableContactMethods.emailAddress
    );

    contactControls.mobilePhoneControl?.setVisible(
      availableContactMethods.mobilePhone
    );
    contactControls.emailControl?.setVisible(
      availableContactMethods.emailAddress
    );
  }

  /**
   * Updates the email field on the case form.
   *
   * This function sets the visibility and requirement level of the email field
   * based on the availability of contact communication channels. If no
   * communication channels (mobile phone or email address) are available, the
   * email field is shown and made mandatory. Otherwise, the email field is
   * hidden and set to optional.
   *
   * @param {Object} formContext  The form context object.
   * @param {Object} availableContactMethods  An object containing the available
   *                                          contact methods.
   * @param {boolean} availableContactMethods.mobilePhone
   * @param {boolean} availableContactMethods.emailAddress
   */
  function _updateCaseEmailField(formContext, availableContactMethods) {
    const isNoContactCommunicationChannelsAvailable =
      !availableContactMethods?.mobilePhone &&
      !availableContactMethods?.emailAddress;

    const requiredLevel = isNoContactCommunicationChannelsAvailable
      ? fieldRequirementLevels.required
      : fieldRequirementLevels.optional;

    formContext
      .getControl(_logicalNames.caseFields.emailAddress)
      .setVisible(isNoContactCommunicationChannelsAvailable);
    formContext
      .getAttribute(_logicalNames.caseFields.emailAddress)
      .setRequiredLevel(requiredLevel);
  }

  /**
   * Returns an object with references to the controls in the contact quick view.
   *
   * @param {Object} formContext - The form context object.
   * @returns {Object} An object containing references to the quick view control,
   *                   mobile phone control, and email control.
   */
  function _getContactViewControls(formContext) {
    const quickViewControl = _tryReadContactQuickViewForm(formContext);

    return {
      quickViewControl,
      mobilePhoneControl: quickViewControl.getControl(
        _logicalNames.contactFields.mobilePhoneNumber
      ),
      emailControl: quickViewControl.getControl(
        _logicalNames.contactFields.emailAddress
      ),
    };
  }

  /**
   * Validates that the execution context is defined and contains a
   * getFormContext method.
   *
   * @param {Object} executionContext - The execution context object.
   * @throws {Error} If the execution context is undefined or does not contain
   *                 a getFormContext method.
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
   * Attempts to read a valid formContext from the execution context and
   * performs validations to ensure the form is associated with the case entity
   * and that both the contact and email address fields are present
   *
   * @param {Object} executionContext  The execution context object containing
   *                                   the form context
   * @returns {Object}  The validated form context
   * @throws {Error}  If the form is not associated with a case entity or if
   *                  required fields are not present
   */
  function _tryReadValidFormContextOrThrow(executionContext) {
    const formContext = executionContext.getFormContext();
    _guardFormIsAssociatedWithCaseEntity(formContext);
    _guardFieldIsPresentInForm(formContext, _logicalNames.caseFields.contact);
    _guardFieldIsPresentInForm(
      formContext,
      _logicalNames.caseFields.emailAddress
    );
    return formContext;
  }

  /**
   * Validates that the form is associated with the case entity.
   * Throws an error if the form is not associated with the case entity.
   *
   * @param {Object} formContext  The form context object
   * @throws {Error}  If the form is not associated with the case entity
   */
  function _guardFormIsAssociatedWithCaseEntity(formContext) {
    if (
      formContext?.contextToken?.entityTypeName !== _logicalNames.tables.case
    ) {
      throw new Error(
        "Invalid form configuration: Form must be associated with" +
          `${_logicalNames.tables.case} entity`
      );
    }
  }

  /**
   * Validates that the specified field is present in the form. Throws an error
   * if the field is not present in the form.
   *
   * @param {Object} formContext  The form context object
   * @param {string} fieldLogicalName  The logical name of the field to check
   * @throws {Error}  If the field is not present in the form
   */
  function _guardFieldIsPresentInForm(formContext, fieldLogicalName) {
    const necessaryField = formContext.getAttribute(fieldLogicalName);
    if (!necessaryField) {
      throw new Error(
        "Invalid form configuration: Form must contain " +
          `${fieldLogicalName} field`
      );
    }
  }

  /**
   * Attempts to read the contact quick view form from the form context.
   *
   * @param {Object} formContext  The form context object.
   * @returns {Object} The quick view control object.
   * @throws {Error} If the quick view control is not found in the form context
   */
  function _tryReadContactQuickViewForm(formContext) {
    const quickViewControl = formContext?.ui?.quickForms.get(
      _logicalNames.controls.contactAvailableMethodsQuickView
    );
    if (!quickViewControl) {
      throw new Error(
        "Invalid form configuration: Form must contain a contact quick view " +
          "form with the name " +
          `"${_logicalNames.controls.contactAvailableMethodsQuickView}"`
      );
    }
    return quickViewControl;
  }

  /**
   * Displays an error dialog to the user with a specified error message and
   * details.
   *
   * @param {Error} error  The error object to display
   */
  function _notifyUserOfError(error) {
    const plugInName = updateAvailableChannelsSection.name;
    _xrm.Navigation.openErrorDialog({
      message: `${plugInName} has encountered an error. ${error.message}`,
      details: error.stack,
    });
  }

  //Return the API
  return {
    updateAvailableChannelsSection,
  };
})();
