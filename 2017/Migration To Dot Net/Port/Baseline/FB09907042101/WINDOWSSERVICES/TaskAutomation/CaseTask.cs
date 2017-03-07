/*
--------------------------------------------------------------------------------------------
Workfile:			CaseTask.cs
Copyright:			Copyright ©2005 Marlborough Stirling

Description:		Implementation of an automation task.
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
PSC		03/10/2005	MAR32 Created
PSC		31/10/2005	MAR117 Correct logging and add comments
PSC		13/12/2005	MAR457 Use omLogging wrapper
PSC		08/06/2006	MAR1859 Add Application Priority
--------------------------------------------------------------------------------------------
*/
using System;
using omLogging = Vertex.Fsd.Omiga.omLogging;  // PSC 13/12/2005 MAR457
using System.Text;

namespace Vertex.Fsd.Omiga.Windows.Service.TaskAutomation
{
	/// <summary>
	/// A task automation case task
	/// </summary>
	internal class CaseTask
	{
		private readonly byte[] _caseActivityGuid;
		private readonly string _stageId;
		private readonly int _caseStageSequenceNo;
		private readonly string _taskId;
		private readonly int _taskInstance;
		private readonly string _applicationNumber;
		private readonly int _applicationFactFindNumber;
		private readonly int _applicationPriority;
		private readonly string _caseActivityGuidAsString;
		private readonly string _activityId;
		private readonly int _activityInstance;
		private readonly string _sourceApplication;

		// PSC 13/12/2005 MAR457
		private static readonly omLogging.Logger _logger = omLogging.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

		#region Constructors

		/// <summary>
		/// Initialises an intance of CaseTask
		/// </summary>
		/// <param name="caseActivityGuid">Case Activity guid</param>
		/// <param name="stageId">Stage Id</param>
		/// <param name="caseStageSequenceNo">Case Stage Sequence Number</param>
		/// <param name="taskId">Task Id</param>
		/// <param name="taskInstance">Task Instance</param>
		/// <param name="applicationNumber">Application Number</param>
		/// <param name="applicationFactFindNumber">Application Fact Find Number</param>
		/// <param name="applicationPriority">Application Priority</param>
		/// <param name="activityId">Activity Id</param>
		/// <param name="activityInstance">Activity Instance</param>
		/// <param name="sourceApplication">Source Application</param>
		public CaseTask(byte[] caseActivityGuid,
			            string stageId,
			            int caseStageSequenceNo,
			            string taskId,
			            int taskInstance,
			            string applicationNumber,
						int applicationFactFindNumber,
						int applicationPriority,
			            string activityId,
			            int activityInstance,
			            string sourceApplication)
		{
			_caseActivityGuid = caseActivityGuid;
			_stageId = stageId;
			_caseStageSequenceNo = caseStageSequenceNo;
			_taskId = taskId;
			_taskInstance = taskInstance;
			_applicationNumber = applicationNumber;
			_applicationFactFindNumber = applicationFactFindNumber;
			_applicationPriority = applicationPriority;	// PSC 08/06/2006 MAR1859
			_activityId = activityId;
			_activityInstance = activityInstance;
			_sourceApplication = sourceApplication;

			StringBuilder thisStringBuilder = new StringBuilder();

			foreach (byte thisByte in _caseActivityGuid)
			{
				thisStringBuilder.Append(thisByte.ToString("X2"));
			}

			_caseActivityGuidAsString = thisStringBuilder.ToString();

		}

		#endregion

		#region Properties

		/// <summary>
		/// <summary>
		/// Property for the Case Activity Guid
		/// </summary>
		/// </summary>
		public byte[] CaseActivityGuid
		{
			get
			{
				return _caseActivityGuid;
			}
		}

		/// <summary>
		/// <summary>
		/// Property for the Stage Id
		/// </summary>
		/// </summary>
		public string StageId
		{
			get
			{
				return _stageId;
			}
		}

		/// <summary>
		/// <summary>
		/// Property for the Case Stage Sequence Number
		/// </summary>
		/// </summary>
		public int CaseStageSequenceNo
		{
			get
			{
				return _caseStageSequenceNo;
			}
		}

		/// <summary>
		/// <summary>
		/// Property for the Task Id
		/// </summary>
		/// </summary>
		public string TaskId
		{
			get
			{
				return _taskId;
			}
		}

		/// <summary>
		/// <summary>
		/// Property for the Task Instance
		/// </summary>
		/// </summary>
		public int TaskInstance
		{
			get
			{
				return _taskInstance;
			}
		}
		
		/// <summary>
		/// <summary>
		/// Property for the Case Activity Guid as a string
		/// </summary>
		/// </summary>
		public string CaseActivityGuidAsString
		{
			get
			{
				return _caseActivityGuidAsString;
			}
		}

		/// <summary>
		/// <summary>
		/// Property for the Application Number
		/// </summary>
		/// </summary>
		public string ApplicationNumber
		{
			get
			{
				return _applicationNumber; 
			}
		}

		/// <summary>
		/// <summary>
		/// Property for the Application Fact Find Number
		/// </summary>
		/// </summary>
		public int ApplicationFactFindNumber
		{
			get
			{
				return _applicationFactFindNumber;
			}
		}

		/// <summary>
		/// <summary>
		/// Property for the Application Priority
		/// </summary>
		/// </summary>
		public int ApplicationPriority
		{
			get
			{
				return _applicationPriority;
			}
		}

		/// <summary>
		/// <summary>
		/// Property for the Activity Guid
		/// </summary>
		/// </summary>
		public string ActivityId
		{
			get
			{
				return _activityId;
			}
		}

		/// <summary>
		/// <summary>
		/// Property for the Activity Instance
		/// </summary>
		/// </summary>
		public int ActivityInstance
		{
			get
			{
				return _activityInstance;
			}
		}

		/// <summary>
		/// <summary>
		/// Property for the Source Application
		/// </summary>
		/// </summary>
		public string SourceApplication
		{
			get
			{
				return _sourceApplication;
			}
		}

		#endregion

		#region Public Methods

		/// <summary>
		/// Produces a string representation of the CaseTask
		/// </summary>
		/// <returns>A string representation of CaseTask</returns>
		public override string ToString()
		{
			return "CaseActivityGuid=" + _caseActivityGuidAsString +
				   " StageId=" + _stageId +
				   " CaseStageSequenceNo=" + _caseStageSequenceNo.ToString() +
				   " TaskId=" + _taskId +
				   " TaskInstance=" + _taskInstance.ToString() +
				   " ApplicationNumber=" + _applicationNumber + 
				   " ApplicationFactFindNumber=" + _applicationFactFindNumber.ToString() +
				   " ApplicationPriority=" + _applicationPriority.ToString(); 
		}

		#endregion
	}
}
