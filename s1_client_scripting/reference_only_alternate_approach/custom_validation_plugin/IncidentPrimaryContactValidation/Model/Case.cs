using System;
using Microsoft.Xrm.Sdk;

namespace IncidentPrimaryContactValidation.Model
{
    internal class Case
    {
        public static readonly string EntityLogicalName = "incident";
        public static readonly string CustomerFieldLogicalName = "customerid";
        public static readonly string ContactFieldLogicalName = "primarycontactid";

        public EntityReference Customer { get; set; }
        public EntityReference Contact { get; set; }

        public Case() { }

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

            Customer = caseEntity.GetAttributeValue<EntityReference>(CustomerFieldLogicalName);
            Contact = caseEntity.GetAttributeValue<EntityReference>(ContactFieldLogicalName);
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
                throw new ArgumentNullException(nameof(entity), "Entity must not be null");
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
