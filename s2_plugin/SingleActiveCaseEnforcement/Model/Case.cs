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
        ///     Initialises a new instance of the <see cref="Case"/> class.
        /// </summary>
        /// <param name="caseEntity">The entity representing the case.</param>
        /// <exception cref="ArgumentNullException">
        ///     Thrown when the case entity is null.
        /// </exception>
        /// <exception cref="ArgumentException">
        ///     Thrown when the case entity does not have the correct logical
        ///     name.
        /// </exception>
        public Case(Entity caseEntity)
        {
            GuardEntityIsNotNull(caseEntity);
            GuardEntityHasTheCorrectType(caseEntity);

            Title = caseEntity.GetAttributeValue<string>(TitleFieldLogicalName);
            Customer = caseEntity.GetAttributeValue<EntityReference>(
                CustomerFieldLogicalName
            );
            Status = caseEntity.GetAttributeValue<OptionSetValue>(
                StatusFieldLogicalName
            );
        }

        /// <summary>
        ///     Explicit casting of an entity to a case
        /// </summary>
        /// <param name="caseEntity">The entity representing the case.</param>
        /// <exception cref="ArgumentNullException">
        ///     Thrown when the case entity is null.
        /// </exception>
        /// <exception cref="ArgumentException">
        ///     Thrown when the case entity does not have the correct logical
        ///     name.
        /// </exception>
        public static explicit operator Case(Entity caseEntity)
        {
            return new Case(caseEntity);
        }

        /// <summary>
        ///     Ensures the entity is not null.
        /// </summary>
        /// <param name="entity">The entity to check.</param>
        /// <exception cref="ArgumentNullException">
        ///     Thrown when the entity is null.
        /// </exception>
        private void GuardEntityIsNotNull(Entity entity)
        {
            if (entity is null)
            {
                throw new ArgumentNullException(
                    nameof(entity),
                    "Entity must not be null"
                );
            }
        }

        /// <summary>
        /// Ensures the entity has the correct logical name.
        /// </summary>
        /// <param name="entity">The entity to check.</param>
        /// <exception cref="ArgumentException">
        ///     Thrown when the entity does not have the correct logical name.
        /// </exception>
        private void GuardEntityHasTheCorrectType(Entity entity)
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
