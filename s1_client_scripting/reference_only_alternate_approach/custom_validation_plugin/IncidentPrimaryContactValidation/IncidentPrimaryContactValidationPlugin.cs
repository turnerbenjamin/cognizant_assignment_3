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

        /// <summary>
        ///     Reads the target value from the plugin context and returns it as
        ///     a Case object.
        /// </summary>
        /// <param name="context">The plugin execution context.</param>
        /// <returns>A Case object representing the target entity.</returns>
        /// <exception cref="ArgumentException">
        ///     Thrown when the target is null.
        /// </exception>
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

        private Case GetPreEntityImage(IPluginExecutionContext context)
        {
            if (context.PreEntityImages.TryGetValue(_preEntityImageAlias, out Entity entity))
            {
                return (Case)entity;
            }
            return new Case();
        }

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

        /// <summary>
        ///     Ensures that the local plugin context is not null.
        /// </summary>
        /// <param name="localPluginContext">
        ///     The local plugin context to check.
        /// </param>
        /// <exception cref="ArgumentNullException">
        ///     Thrown when the local plugin context is null.
        /// </exception>
        private void GuardLocalPluginContextIsNotNull(ILocalPluginContext localPluginContext)
        {
            if (localPluginContext == null)
            {
                throw new ArgumentNullException(nameof(localPluginContext));
            }
        }
    }
}
