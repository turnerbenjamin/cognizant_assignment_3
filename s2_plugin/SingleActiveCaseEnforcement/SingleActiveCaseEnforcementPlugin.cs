using System;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using SingleActiveCaseEnforcement.Model;

namespace SingleActiveCaseEnforcement
{
    public class SingleActiveCaseEnforcementPlugin : PluginBase
    {
        public SingleActiveCaseEnforcementPlugin()
            : base(typeof(SingleActiveCaseEnforcementPlugin)) { }

        /// <summary>
        ///     Plug-in entry point for the Case creation message. Ensures that
        ///     the customer associated with the target case has no other active
        ///     cases.
        /// </summary>
        /// <param name="localPluginContext">The local plugin context.</param>
        /// <exception cref="ArgumentNullException">
        ///     Thrown when the local plugin context is null.
        /// </exception>
        protected override void ExecuteDataversePlugin(
            ILocalPluginContext localPluginContext
        )
        {
            GuardLocalPluginContextIsNotNull(localPluginContext);

            var context = localPluginContext.PluginExecutionContext;
            var service = localPluginContext.PluginUserService;

            var targetCase = ReadTargetCase(context);

            GuardThatCustomerHasNoActiveCases(targetCase, service);
        }

        // Read the target value from the plugin and cast to a case. Throws
        //  if target is null or the cast to case fails
        private Case ReadTargetCase(IPluginExecutionContext context)
        {
            if (
                !context.InputParameters.TryGetValue(
                    "Target",
                    out Entity entity
                )
            )
            {
                throw new ArgumentException("Target is null");
            }

            return Case.CreateFromEntityOrThrow(entity);
        }

        // Check that the customer associated with the target case has no
        // other active cases. If an active case is found, throws an
        // InvalidPluginExecutionException to be displayed to the user.
        private void GuardThatCustomerHasNoActiveCases(
            Case targetCase,
            IOrganizationService service
        )
        {
            var currentlyActiveCaseForCustomer =
                RetrieveActiveCaseForCustomerOrNull(
                    targetCase.Customer.Id,
                    service
                );

            if (currentlyActiveCaseForCustomer != null)
            {
                throw new InvalidPluginExecutionException(
                    "An active case already exists for this customer: "
                        + currentlyActiveCaseForCustomer.Title
                );
            }
        }

        // Look for an active case associated with the given customer ID.
        private Case RetrieveActiveCaseForCustomerOrNull(
            Guid customerId,
            IOrganizationService organizationService
        )
        {
            var query = new QueryExpression(Case.EntityLogicalName)
            {
                ColumnSet = new ColumnSet(Case.TitleFieldLogicalName),
                TopCount = 1,
                Criteria = BuildActiveCasesForCustomerFilter(customerId),
            };

            var result = organizationService.RetrieveMultiple(query);
            return GetFirstCaseFromCollectionOrNull(result);
        }

        // Builds a filter expression to identify active cases associated
        // with the provided customer ID.
        private FilterExpression BuildActiveCasesForCustomerFilter(
            Guid customerId
        )
        {
            const int caseActiveStatusCode = 0;
            var filterExpression = new FilterExpression();
            filterExpression.AddCondition(
                Case.StatusFieldLogicalName,
                ConditionOperator.Equal,
                caseActiveStatusCode
            );

            filterExpression.AddCondition(
                Case.CustomerFieldLogicalName,
                ConditionOperator.Equal,
                customerId
            );
            return filterExpression;
        }

        // Tries to parse the first entity from an EntityCollection as a
        // Case object. If the collection is empty, returns null.
        private Case GetFirstCaseFromCollectionOrNull(EntityCollection result)
        {
            if (result is null || result.Entities.Count == 0)
            {
                return null;
            }
            var caseEntity = result.Entities.First();
            return Case.CreateFromEntityOrThrow(caseEntity);
        }

        // Throws if local plugin context is null.
        private void GuardLocalPluginContextIsNotNull(
            ILocalPluginContext localPluginContext
        )
        {
            if (localPluginContext == null)
            {
                throw new ArgumentNullException(nameof(localPluginContext));
            }
        }
    }
}
