//Declare functionality within a namespace
this.cr4fd = this.window || {};
this.cr4fd.caseFormAvailableCommunicationChannelsUpdate = (function () {
  // eslint-disable-next-line no-undef
  const xrm = Xrm;

  //Dictionary containing logical names for entities and entity fields
  const logicalNames = {
    tables: {
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
  const requirementLevels = {
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
    const formContext = executionContext.getFormContext();
    const contact = await getContactRecord(formContext);

    const availableContactMethods = getContactMethodAvailability(contact);
    updateContactQuickView(formContext, availableContactMethods);
    updateCaseEmailField(formContext, availableContactMethods);
  }

  // Fetch the contact record from the contact lookup field. Returns null if the
  // contact field is not populated
  async function getContactRecord(formContext) {
    const contactFieldValue = formContext
      .getAttribute(logicalNames.caseFields.contact)
      .getValue();
    if (!contactFieldValue || contactFieldValue.length === 0) {
      return null;
    }

    const contactId = contactFieldValue[0].id;

    const contact = await xrm.WebApi.retrieveRecord(
      logicalNames.tables.contact,
      contactId,
      getSelectsQueryStringForContact()
    );
    return contact;
  }

  // Updates the contact quick view. If there are no available communication
  // channeld the view is hidden. Else, a field will be shown within the view
  // for each available contact channel.
  function updateContactQuickView(formContext, availableContactMethods) {
    const contactControls = getContactViewControls(formContext);

    contactControls?.quickViewControl?.setVisible(
      availableContactMethods?.mobilePhone ||
        availableContactMethods?.emailAddress
    );

    contactControls?.mobilePhoneControl?.setVisible(
      availableContactMethods.mobilePhone
    );
    contactControls.emailControl?.setVisible(
      availableContactMethods.emailAddress
    );
  }

  // Updates the email field on the case form. This will be shown and made
  // mandatory where no communication channels can be found for a contact. Else,
  // it is hidden and optional
  function updateCaseEmailField(formContext, availableContactMethods) {
    const isNoContactCommunicationChannelsAvailable =
      !availableContactMethods?.mobilePhone &&
      !availableContactMethods?.emailAddress;

    const requiredLevel = isNoContactCommunicationChannelsAvailable
      ? requirementLevels.required
      : requirementLevels.optional;

    formContext
      .getControl(logicalNames.caseFields.emailAddress)
      .setVisible(isNoContactCommunicationChannelsAvailable);
    formContext
      .getAttribute(logicalNames.caseFields.emailAddress)
      .setRequiredLevel(requiredLevel);
  }

  // Returns an object with references to the controls in the contact quick view
  function getContactViewControls(formContext) {
    const quickViewControl = formContext?.ui?.quickForms.get(
      logicalNames.controls.contactAvailableMethodsQuickView
    );
    const mobilePhoneControl = quickViewControl?.getControl(
      logicalNames.contactFields.mobilePhoneNumber
    );
    const emailControl = quickViewControl?.getControl(
      logicalNames.contactFields.emailAddress
    );

    return {
      quickViewControl,
      mobilePhoneControl,
      emailControl,
    };
  }

  // Returns an object with keys for each channel and bool values indicating
  // whether the channel is available.
  function getContactMethodAvailability(contact) {
    const availability = {
      mobilePhone: false,
      emailAddress: false,
    };

    if (contact) {
      availability.mobilePhone = isContactMethodAvailable(
        contact[logicalNames.contactFields.mobilePhoneNumber],
        contact[logicalNames.contactFields.doNotPhone]
      );
      availability.emailAddress = isContactMethodAvailable(
        contact[logicalNames.contactFields.emailAddress],
        contact[logicalNames.contactFields.doNotEmail]
      );
    }

    return availability;
  }

  // Returns a bool indicating whether a channel is available. Checks that the
  // channel has a value and that the related preferences allow contact through
  // this channel
  function isContactMethodAvailable(channelValue, doNotContactPreference) {
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
  function getSelectsQueryStringForContact(isNested) {
    var prefix = isNested ? "" : "?";
    var fields = [
      logicalNames.contactFields.emailAddress,
      logicalNames.contactFields.doNotEmail,
      logicalNames.contactFields.mobilePhoneNumber,
      logicalNames.contactFields.doNotPhone,
    ];
    return prefix + "$select=" + fields.join(",");
  }

  //Return the API
  return {
    updateAvailableChannelsSection,
  };
})();
