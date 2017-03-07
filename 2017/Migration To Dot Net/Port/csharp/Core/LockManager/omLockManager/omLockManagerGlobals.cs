using System;
using Vertex.Fsd.Omiga.VisualBasicPort; 
namespace omLockManager
{
    public class omLockManagerGlobals
    {

        // If this is declared in StdData.bas then all other components have to be rebuilt.
        public const string gstrLOCKMANAGER_COMPONENT = "omLockManager";
        public static void Main() 
        {
            adoAssistEx.adoBuildDbConnectionString();
        }


    }

}
