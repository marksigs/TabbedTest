using System;
//using COMSVCSLib;
using System.Transactions;
using System.Runtime.InteropServices;
using System.EnterpriseServices;
namespace omLockManager
{
    //public class LockManagerNTxBO : COMSVCSLib.ObjectControl
    public class LockManagerNTxBO : ServicedComponent
    {

        // Workfile:      LockManagerNTxBO.cls
        // Copyright:     Copyright © 2005 Marlborough Stirling
        // Description:   A transactional manager object that co-ordinates the calls to create and
        // release locks for the customer and application.
        // Note: XML is not used internally within this module, but only when interfacing
        // to other components that have an XML based interface.
        // Dependencies:
        // ------------------------------------------------------------------------------------------
        // History:
        // 
        // Prog   Date        Description
        // AS     05/04/2005  Created
        // ------------------------------------------------------------------------------------------
        private COMSVCSLib.ObjectContext m_objContext = null;
        public string omRequest(string vstrXmlIn) 
        {
            LockManagerBO objLockManagerBO;
            // WARNING: On Error GOTO omRequestExit is not supported
            try 
            {
                if (m_objContext == null)
                {
                    objLockManagerBO = new LockManagerBO();
                }
                else
                {
                    objLockManagerBO = m_objContext.CreateInstance(omLockManagerGlobals.gstrLOCKMANAGER_COMPONENT + ".LockManagerBO");
                }
                // Delegate to BO.
                // WARNING: omRequestExit: is not supported 
            }
            catch(Exception exc)
            {
                objLockManagerBO = null;
            }
            return objLockManagerBO.omRequest(vstrXmlIn);
        }

        private void ObjectControl_Activate() 
        {
            object GetObjectContext = null;
            m_objContext = GetObjectContext;
        }

        private bool ObjectControl_CanBePooled() 
        {
            return false;
        }

        private void ObjectControl_Deactivate() 
        {
            m_objContext = null;
        }


    }

}
