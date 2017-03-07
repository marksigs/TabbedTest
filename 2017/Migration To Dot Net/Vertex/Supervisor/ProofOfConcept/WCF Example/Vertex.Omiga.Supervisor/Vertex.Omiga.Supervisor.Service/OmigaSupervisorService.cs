using System;
using System.Collections.Generic;
using System.ServiceModel;
using System.Runtime.Serialization;

using Vertex.Omiga.Supervisor.Business;
using Vertex.Omiga.Supervisor.Service.Interface;
using Vertex.Omiga.Supervisor.Service.Types.Responses;
using Vertex.Omiga.Supervisor.Service.Types.Requests;
using Vertex.Omiga.Supervisor.Service.Types.Faults;
using Vertex.Omiga.Supervisor.Service.Types.Entities;

using BusinessEntities = Vertex.Omiga.Supervisor.Business.Entities;
using ServiceEntities = Vertex.Omiga.Supervisor.Service.Types.Entities;

namespace Vertex.Omiga.Supervisor.Service
{
    /// <summary>
    /// Omiga Supervisor service implementation
    /// </summary>
    [ServiceBehavior(IncludeExceptionDetailInFaults = true)]
    public class OmigaSupervisorService : BaseService, IOmigaSupervisorService
    {
        /// <summary>
        /// Retrieve all <see cref="GlobalParameter"/> entities
        /// </summary>
        /// <param name="request">Request</param>
        /// <returns>Response</returns>
        GetGlobalParametersResponse IOmigaSupervisorService.GetGlobalParameters(GetGlobalParametersRequest request)
        {
            try
            {
                PrepareAuthenticatedRequest(request);

                // call core business logic
                //todo: pass paging and sorting info to DAL
                BusinessEntities.GlobalParameter[] businessEntities =
                    GlobalParameters.GetGlobalParameters();

                // map business entities to service entities
                ServiceEntities.GlobalParameter[] serviceEntities =
                    GlobalParameterMapping.CreateFromBusinessEntities(businessEntities);

                // create, populate and return response
                GetGlobalParametersResponse response = new GetGlobalParametersResponse(request.Context);
                response.GlobalParameters = serviceEntities;

                PrepareResponse(response);

                return response;
            }

            catch (Exception ex)
            {
                ThrowServiceFault(ex);
                throw; // to keep the compiler happy
            }
        }

        /// <summary>
        /// Create a new <see cref="GlobalParameter"/> entity
        /// </summary>
        /// <param name="request">Request</param>
        /// <returns>Response</returns>
        CreateGlobalParameterResponse IOmigaSupervisorService.CreateGlobalParameter(CreateGlobalParameterRequest request)
        {
            try
            {
                PrepareAuthenticatedRequest(request);

                // map service entities to business entities
                BusinessEntities.GlobalParameter businessEntity =
                    GlobalParameterMapping.CreateFromServiceEntity(request.EntityToCreate);

                // call core business logic
                BusinessEntities.GlobalParameter newBusinessEntity =
                    GlobalParameters.CreateGlobalParameter(businessEntity);

                // map new business entity to service entity
                ServiceEntities.GlobalParameter serviceEntity =
                    GlobalParameterMapping.CreateFromBusinessEntity(newBusinessEntity);

                // create, populate and return response
                CreateGlobalParameterResponse response = new CreateGlobalParameterResponse(request.Context);
                response.NewEntity = serviceEntity;

                PrepareResponse(response);

                return response;
            }

            catch (Exception ex)
            {
                ThrowServiceFault(ex);
                throw; // to keep the compiler happy
            }
        }

        /// <summary>
        /// Update an existing <see cref="GlobalParameter"/> entity
        /// </summary>
        /// <param name="request">Request</param>
        /// <returns>Response</returns>
        UpdateGlobalParameterResponse IOmigaSupervisorService.UpdateGlobalParameter(UpdateGlobalParameterRequest request)
        {
            try
            {
                PrepareAuthenticatedRequest(request);

                // map service entities to business entities
                BusinessEntities.GlobalParameter businessEntity =
                    GlobalParameterMapping.CreateFromServiceEntity(request.EntityToUpdate);

                // call core business logic
                BusinessEntities.GlobalParameter updatedBusinessEntity =
                    GlobalParameters.UpdateGlobalParameter(businessEntity);

                // map updated business entity to service entity
                ServiceEntities.GlobalParameter serviceEntity =
                    GlobalParameterMapping.CreateFromBusinessEntity(updatedBusinessEntity);

                // create, populate and return response
                UpdateGlobalParameterResponse response = new UpdateGlobalParameterResponse(request.Context);
                response.UpdatedEntity = serviceEntity;

                PrepareResponse(response);

                return response;
            }

            catch (Exception ex)
            {
                ThrowServiceFault(ex);
                throw; // to keep the compiler happy
            }
        }

        /// <summary>
        /// Delete an existing <see cref="GlobalParameter"/> entity
        /// </summary>
        /// <param name="request">Request</param>
        /// <returns>Response</returns>
        DeleteGlobalParameterResponse IOmigaSupervisorService.DeleteGlobalParameter(DeleteGlobalParameterRequest request)
        {
            try
            {
                PrepareAuthenticatedRequest(request);

                // map service entities to business entities
                BusinessEntities.GlobalParameterIdentity businessEntity =
                    GlobalParameterMapping.CreateFromServiceEntity(request.EntityToDelete);

                // call core business logic
                GlobalParameters.DeleteGlobalParameter(businessEntity);

                // create, populate and return response
                DeleteGlobalParameterResponse response = new DeleteGlobalParameterResponse(request.Context);

                PrepareResponse(response);

                return response;
            }

            catch (Exception ex)
            {
                ThrowServiceFault(ex);
                throw; // to keep the compiler happy
            }
        }

    }

}
