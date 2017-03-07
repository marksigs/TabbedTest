using System;

using Vertex.Omiga.Supervisor.Service.Types.Contexts;
using Vertex.Omiga.Supervisor.Service.Types.Requests;

namespace Vertex.Omiga.Supervisor.Client
{
    /// <summary>
    /// Utility class for service clients
    /// </summary>
    internal static class ClientUtility
    {
        /// <summary>
        /// Populates the request context provided
        /// </summary>
        /// <param name="context">Context to populate</param>
        internal static void PopulateRequestContext(RequestContext context)
        {
            PopulateRequestContext(context, Guid.NewGuid());
        }

        /// <summary>
        /// Populates the provided request context including an existing 
        /// authentication token
        /// </summary>
        /// <param name="context">Context to populate</param>
        /// <param name="authenticationToken">Existing authentication 
        /// token</param>
        internal static void PopulateRequestContext(RequestContext context, Guid authenticationToken)
        {
            context.ContextGUID = Guid.NewGuid();
            context.AuthenticationToken = authenticationToken;
            context.RequestSource = "Vertex.Omiga.Supervisor.Client";
        }

    }

}
