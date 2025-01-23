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

        /// <summary>
        /// Initialises a new instance of the <see cref="Case"/> class.
        /// </summary>
        /// <param name="title">The title of the case</param>
        /// <param name="customer">The customer associated with the case</param>
        /// <param name="status">The status of the case</param>
        private Case(
            string title,
            EntityReference customer,
            OptionSetValue status
        )
        {
            Title = title;
            Customer = customer;
            Status = status;
        }

        /// <summary>
        /// Creates a new Case instance from the provided Entity object.
        /// </summary>
        /// <param name="caseEntity">
        /// The Entity object representing a case.
        /// </param>
        /// <returns>
        /// A new Case instance initialized with the entity's attributes.
        /// </returns>
        /// <exception cref="ArgumentNullException">
        /// Thrown when the entity is null.
        /// </exception>
        /// <exception cref="ArgumentException">
        /// Thrown when the entity does not have the correct logical name.
        /// </exception>
        public static Case CreateFromEntityOrThrow(Entity caseEntity)
        {
            GuardEntityIsNotNull(caseEntity);
            GuardEntityHasTheCorrectType(caseEntity);

            var title = caseEntity.GetAttributeValue<string>(
                TitleFieldLogicalName
            );
            var customer = caseEntity.GetAttributeValue<EntityReference>(
                CustomerFieldLogicalName
            );
            var status = caseEntity.GetAttributeValue<OptionSetValue>(
                StatusFieldLogicalName
            );

            return new Case(title, customer, status);
        }

        // Ensures the entity is not null.
        private static void GuardEntityIsNotNull(Entity entity)
        {
            if (entity is null)
            {
                throw new ArgumentNullException(
                    nameof(entity),
                    "Entity must not be null"
                );
            }
        }

        // Ensures the entity has the correct logical name.
        private static void GuardEntityHasTheCorrectType(Entity entity)
        {
            if (entity.LogicalName != EntityLogicalName)
            {
                throw new ArgumentException(
                    $"Entity must be of type {EntityLogicalName}",
                    nameof(entity)
                );
            }
        }
    }
}
