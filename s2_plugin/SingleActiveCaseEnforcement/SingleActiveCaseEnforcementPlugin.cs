using System;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using SingleActiveCaseEnforcement.Model;

namespace SingleActiveCaseEnforcement
{
    public class SingleActiveCaseEnforcenentPlugin : PluginBase
    {
        public SingleActiveCaseEnforcenentPlugin(
            string unsecureConfiguration,
            string secureConfiguration
        )
            : base(typeof(SingleActiveCaseEnforcenentPlugin)) { }

        protected override void ExecuteDataversePlugin(
            ILocalPluginContext localPluginContext
        )
        {
            if (localPluginContext == null)
            {
                throw new ArgumentNullException(nameof(localPluginContext));
            }

            var context = localPluginContext.PluginExecutionContext;
            var tracingService = localPluginContext.TracingService;
            var service = localPluginContext.PluginUserService;

            var targetCase = GetTargetCase(context);
            var currentlyActiveCaseForCustomer = FindActiveCaseForCustomer(
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
            tracingService.Trace(
                $"Case creation permitted: {targetCase.Title}"
            );
        }

        private Case FindActiveCaseForCustomer(
            Guid customerId,
            IOrganizationService organizationService
        )
        {
            var query = new QueryExpression(Case.EntityLogicalName)
            {
                ColumnSet = new ColumnSet(Case.TitleFieldLogicalName),
                TopCount = 1,
            };

            const int caseActiveStatusCode = 0;
            query.Criteria.AddCondition(
                Case.StatusFieldLogicalName,
                ConditionOperator.Equal,
                caseActiveStatusCode
            );

            query.Criteria.AddCondition(
                Case.CustomerFieldLogicalName,
                ConditionOperator.Equal,
                customerId
            );

            var result = organizationService.RetrieveMultiple(query);
            var entity = result.Entities.FirstOrDefault<Entity>();
            if (entity is null)
            {
                return null;
            }
            return new Case(entity);
        }

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
            return new Case(entity);
        }
    }
}
