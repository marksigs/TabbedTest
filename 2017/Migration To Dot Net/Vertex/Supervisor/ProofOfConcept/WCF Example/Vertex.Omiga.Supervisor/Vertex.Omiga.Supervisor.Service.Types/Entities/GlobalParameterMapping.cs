using System;
using System.Collections.Generic;

using BusinessTypes = Vertex.Omiga.Supervisor.Business.Entities;
using ServiceTypes = Vertex.Omiga.Supervisor.Service.Types.Entities;

namespace Vertex.Omiga.Supervisor.Service.Types.Entities
{
    /// <summary>
    /// Service / business entity mapping class
    /// </summary>
    public static class GlobalParameterMapping
    {
        /// <summary>
        /// Create a service entity array from a business entity array.
        /// Entity detail
        /// </summary>
        /// <param name="businessEntities">Business entity array</param>
        /// <returns>New service entity array</returns>
        public static ServiceTypes.GlobalParameter[] CreateFromBusinessEntities(BusinessTypes.GlobalParameter[] businessEntities)
        {
            List<ServiceTypes.GlobalParameter> serviceEntities = new List<GlobalParameter>();

            foreach (BusinessTypes.GlobalParameter businessEntity in businessEntities)
            {
                ServiceTypes.GlobalParameter serviceEntity = new ServiceTypes.GlobalParameter();

                // map business entity to service entity
                serviceEntity.TimeStamp = businessEntity.TimeStamp;
                serviceEntity.Name = businessEntity.Name;
                serviceEntity.Description = businessEntity.Description;
                serviceEntity.StartDate = businessEntity.StartDate;
                serviceEntity.ValueString = businessEntity.ValueString;
                serviceEntity.ValueAmount = businessEntity.ValueAmount;
                serviceEntity.ValueMaximumAmount = businessEntity.ValueMaximumAmount;
                serviceEntity.ValuePercentage = businessEntity.ValuePercentage;
                serviceEntity.ValueBoolean = businessEntity.ValueBoolean;

                serviceEntities.Add(serviceEntity);
            }

            return serviceEntities.ToArray();
        }

        /// <summary>
        /// Create a service entity from a business entity.
        /// Entity detail
        /// </summary>
        /// <param name="businessEntity">Business entity</param>
        /// <returns>New service entity instance</returns>
        public static ServiceTypes.GlobalParameter CreateFromBusinessEntity(BusinessTypes.GlobalParameter businessEntity)
        {
            ServiceTypes.GlobalParameter[] serviceEntities =
                CreateFromBusinessEntities(new BusinessTypes.GlobalParameter[] { businessEntity });

            return serviceEntities[0];
        }

        /// <summary>
        /// Create a business entity from a service entity.
        /// Entity detail
        /// </summary>
        /// <param name="serviceEntity">Service entity</param>
        /// <returns>New business entity instance</returns>
        public static BusinessTypes.GlobalParameter CreateFromServiceEntity(ServiceTypes.GlobalParameter serviceEntity)
        {
            BusinessTypes.GlobalParameter businessEntity = new BusinessTypes.GlobalParameter();

            // map service entity to business entity
            businessEntity.TimeStamp = serviceEntity.TimeStamp;
            businessEntity.Name = serviceEntity.Name;
            businessEntity.Description = serviceEntity.Description;
            businessEntity.StartDate = serviceEntity.StartDate;
            businessEntity.ValueString = serviceEntity.ValueString;
            businessEntity.ValueAmount = serviceEntity.ValueAmount;
            businessEntity.ValueMaximumAmount = serviceEntity.ValueMaximumAmount;
            businessEntity.ValuePercentage = serviceEntity.ValuePercentage;
            businessEntity.ValueBoolean = serviceEntity.ValueBoolean;

            return businessEntity;
        }

        /// <summary>
        /// Create a business entity from a service entity.
        /// Identity detail
        /// </summary>
        /// <param name="serviceEntity">Service entity</param>
        /// <returns>New business entity instance</returns>
        public static BusinessTypes.GlobalParameterIdentity CreateFromServiceEntity(ServiceTypes.GlobalParameterIdentity serviceEntity)
        {
            BusinessTypes.GlobalParameterIdentity businessEntity = new BusinessTypes.GlobalParameterIdentity();

            // map service entity to business entity
            businessEntity.TimeStamp = serviceEntity.TimeStamp;
            businessEntity.Name = serviceEntity.Name;
            businessEntity.StartDate = serviceEntity.StartDate;

            return businessEntity;
        }

    }

}
