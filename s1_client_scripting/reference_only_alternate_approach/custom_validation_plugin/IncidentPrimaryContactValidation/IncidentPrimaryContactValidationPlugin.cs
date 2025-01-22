using System;
using IncidentPrimaryContactValidation.Model;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace IncidentPrimaryContactValidation
{
    public class IncidentPrimaryContactValidationPlugin : PluginBase
    {
        private const string _accountLogicalName = "account";
        private const string _contactLogicalName = "contact";
        private const string _associatedAccountFieldLogicalName = "parentcustomerid";
        private const string _preEntityImageAlias = "IncidentContactAndCustomerValues";

        public IncidentPrimaryContactValidationPlugin()
            : base(typeof(IncidentPrimaryContactValidationPlugin)) { }

        // Entry point for custom business logic execution
        protected override void ExecuteDataversePlugin(ILocalPluginContext localPluginContext)
        {
            if (localPluginContext == null)
            {
                throw new ArgumentNullException(nameof(localPluginContext));
            }
            GuardLocalPluginContextIsNotNull(localPluginContext);

            var context = localPluginContext.PluginExecutionContext;
            var service = localPluginContext.PluginUserService;

            var preEntityImage = GetPreEntityImage(context);
            var caseAfterCreateOrUpdate = GetCaseAfterCreateOrUpdate(preEntityImage, context);

            ValidateContactField(caseAfterCreateOrUpdate, service);
        }

        // Update the preEntityImage with updates from context so that it
        // reflects the post create/update state
        private Case GetCaseAfterCreateOrUpdate(
            Case preEntityImage,
            IPluginExecutionContext context
        )
        {
            if (!context.InputParameters.TryGetValue("Target", out Entity entity))
            {
                throw new ArgumentException("Target is null");
            }

            if (entity.Attributes.ContainsKey(Case.ContactFieldLogicalName))
            {
                preEntityImage.Contact =
                    entity.Attributes[Case.ContactFieldLogicalName] as EntityReference;
            }

            if (entity.Attributes.ContainsKey(Case.CustomerFieldLogicalName))
            {
                preEntityImage.Customer =
                    entity.Attributes[Case.CustomerFieldLogicalName] as EntityReference;
            }

            return preEntityImage;
        }

        // Fetch the pre entity image and cast to case. If no image exists
        // a case is returned with null properties
        private Case GetPreEntityImage(IPluginExecutionContext context)
        {
            if (context.PreEntityImages.TryGetValue(_preEntityImageAlias, out Entity entity))
            {
                return (Case)entity;
            }
            return new Case();
        }

        // Ensures that the value in the contact field is associated with the
        // value in the customer field where both fields have been set and the
        // record in the customer field is an account.
        private void ValidateContactField(Case targetCase, IOrganizationService service)
        {
            // Retrieve the contact and customer values from the Case entity
            var contactValue = targetCase.Contact;
            var customerValue = targetCase.Customer;

            // Return early if contact and customer fields are not both set and
            // if the customer field type is not an account
            if (
                contactValue == null
                || customerValue == null
                || customerValue.LogicalName != _accountLogicalName
            )
            {
                return;
            }

            // Fetch the customer record
            var contact = service.Retrieve(
                _contactLogicalName,
                contactValue.Id,
                new ColumnSet(_associatedAccountFieldLogicalName)
            );

            // Read the customer's parent account field
            var contactParentAccountValue = contact.GetAttributeValue<EntityReference>(
                _associatedAccountFieldLogicalName
            );

            if (
                contactParentAccountValue == null
                || contactParentAccountValue.Id != customerValue.Id
            )
            {
                throw new InvalidPluginExecutionException(
                    "The contact is not associated with the customer."
                );
            }
        }

        // Throws an error if the localPluginContext is null
        private void GuardLocalPluginContextIsNotNull(ILocalPluginContext localPluginContext)
        {
            if (localPluginContext == null)
            {
                throw new ArgumentNullException(nameof(localPluginContext));
            }
        }
    }
}
