using System;
using System.Diagnostics;
using System.Text;
using System.ServiceModel;

using Vertex.Omiga.Supervisor.Service.Types.Enums;
using Vertex.Omiga.Supervisor.Service.Types.Faults;
using Vertex.Omiga.Supervisor.Service.Types.Requests;
using Vertex.Omiga.Supervisor.Service.Types.Responses;

namespace Vertex.Omiga.Supervisor.Service
{
    /// <summary>
    /// Base class for service implementations
    /// </summary>
    public abstract class BaseService
    {
        /// <summary>
        /// Prepare a request to be processed by the service. 
        /// Includes validating requesting application.
        /// Should only be called for service calls that do NOT require a
        /// valid authentication token
        /// </summary>
        /// <param name="request">Service request</param>
        protected static void PrepareRequest(BaseOperationRequest request)
        {
            #region Validate arguments
            if (request == null) 
                throw new ArgumentNullException("request");

            request.Validate();
            
            #endregion

            // stamp request context
            request.Context.ReceivedAt = DateTime.Now;

            // authenticate request and check authorisation
            ValidateRequestSource(request);
        }

        /// <summary>
        /// Prepare a request to be processed by the service.
        /// Includes authenticating the security token, validating the 
        /// requesting application and checking user authorisation to call
        /// the specified service operation
        /// </summary>
        /// <param name="request">Service request</param>
        protected static void PrepareAuthenticatedRequest(BaseOperationRequest request)
        {
            #region Validate arguments
            if (request == null) 
                throw new ArgumentNullException("request");

            #endregion

            PrepareRequest(request);

            // additionally, authenticate the security token
            ValidateAuthenticationToken(request);

            // ensure the calling user can access this service operation
            ValidateOperationAuthority(request, request.GetType().FullName);
        }

        /// <summary>
        /// Prepare a response to be returned from the service
        /// </summary>
        /// <param name="response">Service response</param>
        protected static void PrepareResponse(BaseOperationResponse response)
        {
            #region Validate arguments
            if (response == null) 
                throw new ArgumentNullException("response");

            response.Validate();

            #endregion

            // stamp response context
            response.Context.SentAt = DateTime.Now;
        }

        /// <summary>
        /// Records an exception and rethrows as a service fault
        /// </summary>
        /// <param name="exception">Exception</param>
        protected static void ThrowServiceFault(Exception exception)
        {
            // store exception details
            Guid faultGuid = Guid.NewGuid();

            // store detailed fault information in event log etc.
            StoreExceptionInfo(faultGuid, exception);

            // rethrow as a service fault
            throw new FaultException<OmigaSupervisorFault>(new OmigaSupervisorFault(exception.Message, faultGuid), "Service error");
        }

        #region Helpers
        /// <summary>
        /// Perform request source validation
        /// </summary>
        /// <param name="request">Service request</param>
        private static void ValidateRequestSource(BaseOperationRequest request)
        {
            //todo: example code only. implement better request validation solution
            Console.WriteLine("Validate {0}", request.Context.RequestSource);

            if (request.Context.RequestSource != "Vertex.Omiga.Supervisor.Client" &&
                request.Context.RequestSource != "Vertex.Omiga.Supervisor.Web")
                throw new Exception("Unrecognised request source");
        }

        /// <summary>
        /// Perform request authentication
        /// </summary>
        /// <param name="request">Service request</param>
        private static void ValidateAuthenticationToken(BaseOperationRequest request)
        {
            //todo: example code only. implement better user authentication solution
            Console.WriteLine("Authenticate {0}", request.Context.AuthenticationToken);

            if (request.Context.AuthenticationToken == Guid.Empty)
                throw new Exception("Invalid authentication token");
        }

        /// <summary>
        /// Perform request authentication
        /// </summary>
        /// <param name="request">Service request</param>
        /// <param name="operation">Service operation name</param>
        private static void ValidateOperationAuthority(BaseOperationRequest request, string operation)
        {
            //todo: implement user/operation authorisation check
            Console.WriteLine("Authorise {0}", operation);
        }

        /// <summary>
        /// Record exception information to the event log
        /// </summary>
        /// <param name="faultGuid">Unique ID for this exception</param>
        /// <param name="exception">Exception</param>
        private static void StoreExceptionInfo(Guid faultGuid, Exception exception)
        {
            // store exception details
            string exceptionInfo = GetExceptionInfo(faultGuid, exception);

            EventLog.WriteEntry("Vertex Omiga Supervisor Service", exceptionInfo, EventLogEntryType.Error);
        }

        /// <summary>
        /// Builds a string describing the exception which is suitable for
        /// logging
        /// </summary>
        /// <param name="faultGuid">Unique ID for this exception</param>
        /// <param name="exception">Exception</param>
        /// <returns>String</returns>
        private static string GetExceptionInfo(Guid faultGuid, Exception exception)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendFormat("Fault ID: {0}", faultGuid);
            sb.AppendLine();

            Exception currentException = exception;

            while (currentException != null)
            {
                sb.AppendLine("==> " + currentException.Message);
                sb.AppendLine(currentException.StackTrace);

                currentException = currentException.InnerException;

                if (currentException != null)
                {
                    sb.AppendLine();
                    sb.AppendLine();
                }
            }

            return sb.ToString();
        }

        #endregion

    }

}
