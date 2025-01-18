using System;
using Microsoft.Xrm.Sdk;

namespace SingleActiveCaseEnforcement.Model
{
    internal class Case
    {
        public static readonly string EntityLogicalName = "incident";
        public static readonly string TitleFieldLogicalName = "title";
        public static readonly string CustomerFieldLogicalName = "customerid";
        public static readonly string StatusFieldLogicalName = "statecode";

        public string Title { get; }
        public EntityReference Customer { get; }
        public OptionSetValue Status { get; }

        public Case(Entity caseEntity)
        {
            if (caseEntity is null)
            {
                throw new ArgumentNullException(
                    $"{nameof(caseEntity)} must not be null"
                );
            }

            if (caseEntity.LogicalName != EntityLogicalName)
            {
                throw new ArgumentException(
                    $"Target must be of type {EntityLogicalName}"
                );
            }

            Title = caseEntity.GetAttributeValue<string>(TitleFieldLogicalName);
            Customer = caseEntity.GetAttributeValue<EntityReference>(
                CustomerFieldLogicalName
            );
            Status = caseEntity.GetAttributeValue<OptionSetValue>(
                StatusFieldLogicalName
            );
        }
    }
}
