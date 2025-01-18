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

        /// <summary>
        ///     Reads the target value from the plugin context and returns it as
        ///     a Case object.
        /// </summary>
        /// <param name="context">The plugin execution context.</param>
        /// <returns>A Case object representing the target entity.</returns>
        /// <exception cref="ArgumentException">
        ///     Thrown when the target is null.
        /// </exception>
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

            return (Case)entity;
        }

        /// <summary>
        ///     Ensures that the customer associated with the target case has no
        ///     other active cases. If an active case is found, throws an
        ///     InvalidPluginExecutionException to be displayed to the user.
        /// </summary>
        /// <param name="targetCase">The target case to check.</param>
        /// <param name="service">
        ///     The organisation service to fetch cases from the Dataverse.
        /// </param>
        /// <exception cref="InvalidPluginExecutionException">
        ///     Thrown when an active case is found for the customer.
        /// </exception>
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

        /// <summary>
        ///     Looks for an active case associated with the given customer ID.
        /// </summary>
        /// <param name="customerId">
        ///     The ID of the customer to check for active cases.
        /// </param>
        /// <param name="organizationService">
        ///     The organization service to use for the query.
        /// </param>
        /// <returns>
        ///     A Case object if an active case is found; otherwise, null.
        /// </returns>
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

        /// <summary>
        ///     Builds a filter expression to identify active cases associated
        ///     with the provided customer ID.
        /// </summary>
        /// <param name="customerId">
        ///     The ID of the customer to check for active cases.
        /// </param>
        /// <returns>A FilterExpression object to be used in a query.</returns>
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

        /// <summary>
        ///     Tries to parse the first entity from an EntityCollection as a
        ///     Case object. If the collection is empty, returns null.
        /// </summary>
        /// <param name="result">The EntityCollection to parse.</param>
        /// <returns>
        ///     A Case object if an entity is found; otherwise, null.
        /// </returns>
        private Case GetFirstCaseFromCollectionOrNull(EntityCollection result)
        {
            if (result is null || result.Entities.Count == 0)
            {
                return null;
            }
            return (Case)result.Entities.First();
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
