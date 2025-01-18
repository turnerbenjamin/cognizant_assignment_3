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
        /// Plug-in entry point to be used with the Case creation message.
        /// Throws an error if the customer associated with a target case is
        /// associated with one or more active cases.
        /// </summary>
        /// <param name="localPluginContext">The local plugin context.</param>
        protected override void ExecuteDataversePlugin(
            ILocalPluginContext localPluginContext
        )
        {
            if (localPluginContext == null)
            {
                throw new ArgumentNullException(nameof(localPluginContext));
            }

            var context = localPluginContext.PluginExecutionContext;
            var service = localPluginContext.PluginUserService;

            var targetCase = GetTargetCase(context);

            GuardThatCustomerHasNoActiveCases(targetCase, service);
        }

        // Looks for an active case associated with the same customer. If an
        // active case is found, throws an Invalid PluginExecutionException to
        // be displayed to the user
        private void GuardThatCustomerHasNoActiveCases(
            Case targetCase,
            IOrganizationService service
        )
        {
            var currentlyActiveCaseForCustomer = FindActiveCaseForCustomer(
                targetCase.Customer.Id,
                service
            );

            if (currentlyActiveCaseForCustomer is null)
            {
                return;
            }

            throw new InvalidPluginExecutionException(
                "An active case already exists for this customer: "
                    + currentlyActiveCaseForCustomer.Title
            );
        }

        // Looks for an active case associated with the given customer id.
        // Returns a case where one is found. Else, returns null.
        private Case FindActiveCaseForCustomer(
            Guid customerId,
            IOrganizationService organizationService
        )
        {
            var query = new QueryExpression(Case.EntityLogicalName)
            {
                ColumnSet = new ColumnSet(Case.TitleFieldLogicalName),
                TopCount = 1,
                Criteria = GetActiveCasesForCustomerFilterExpression(
                    customerId
                ),
            };

            var result = organizationService.RetrieveMultiple(query);
            return ParseFirstEntityAsCaseFromEntityCollection(result);
        }

        // Returns a filter expression used to identify active cases associated
        // with the provided customer id
        private FilterExpression GetActiveCasesForCustomerFilterExpression(
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

        // Tries to parse the first entity from a case collection as a case. If
        // the collection is empty, returns null
        private Case ParseFirstEntityAsCaseFromEntityCollection(
            EntityCollection result
        )
        {
            if (result is null || result.Entities.Count == 0)
            {
                return null;
            }
            return (Case)result.Entities.First();
        }

        // Reads the target value from the plugin context and returns this as a
        // case. Throws an error if target is null
        private Case GetTargetCase(IPluginExecutionContext context)
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

            return (Case)entity;
        }
    }
}
