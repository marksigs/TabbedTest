
#pragma warning disable 162

namespace MSC.IntegrationService
{

    [Microsoft.XLANGs.BaseTypes.PortTypeOperationAttribute(
        "ProcessRequest",
        new System.Type[]{
            typeof(MSC.IntegrationService.__messagetype_MSC_IntegrationService_Schemas_cdiInterop), 
            typeof(MSC.IntegrationService.__messagetype_System_Xml_XmlDocument)
        },
        new string[]{
        }
    )]
    [Microsoft.XLANGs.BaseTypes.PortTypeAttribute(Microsoft.XLANGs.BaseTypes.EXLangSAccess.ePublic, "")]
    [System.SerializableAttribute]
    sealed public class MSCReceive_PortType : Microsoft.BizTalk.XLANGs.BTXEngine.BTXPortBase
    {
        public MSCReceive_PortType(int portInfo, Microsoft.XLANGs.Core.IServiceProxy s)
            : base(portInfo, s)
        { }
        public MSCReceive_PortType(MSCReceive_PortType p)
            : base(p)
        { }

        public override Microsoft.XLANGs.Core.PortBase Clone()
        {
            MSCReceive_PortType p = new MSCReceive_PortType(this);
            return p;
        }

        public static readonly Microsoft.XLANGs.BaseTypes.EXLangSAccess __access = Microsoft.XLANGs.BaseTypes.EXLangSAccess.ePublic;
        #region port reflection support
        static public Microsoft.XLANGs.Core.OperationInfo ProcessRequest = new Microsoft.XLANGs.Core.OperationInfo
        (
            "ProcessRequest",
            System.Web.Services.Description.OperationFlow.RequestResponse,
            typeof(MSCReceive_PortType),
            typeof(__messagetype_MSC_IntegrationService_Schemas_cdiInterop),
            typeof(__messagetype_System_Xml_XmlDocument),
            new System.Type[]{},
            new string[]{}
        );
        static public System.Collections.Hashtable OperationsInformation
        {
            get
            {
                System.Collections.Hashtable h = new System.Collections.Hashtable();
                h[ "ProcessRequest" ] = ProcessRequest;
                return h;
            }
        }
        #endregion // port reflection support
    }

    [Microsoft.XLANGs.BaseTypes.PortTypeOperationAttribute(
        "SendToMsgBox",
        new System.Type[]{
            typeof(MSC.IntegrationService.__messagetype_System_Xml_XmlDocument), 
            typeof(MSC.IntegrationService.__messagetype_System_Xml_XmlDocument)
        },
        new string[]{
        }
    )]
    [Microsoft.XLANGs.BaseTypes.PortTypeAttribute(Microsoft.XLANGs.BaseTypes.EXLangSAccess.ePublic, "")]
    [System.SerializableAttribute]
    sealed public class SendRequest_PortType : Microsoft.BizTalk.XLANGs.BTXEngine.BTXPortBase
    {
        public SendRequest_PortType(int portInfo, Microsoft.XLANGs.Core.IServiceProxy s)
            : base(portInfo, s)
        { }
        public SendRequest_PortType(SendRequest_PortType p)
            : base(p)
        { }

        public override Microsoft.XLANGs.Core.PortBase Clone()
        {
            SendRequest_PortType p = new SendRequest_PortType(this);
            return p;
        }

        public static readonly Microsoft.XLANGs.BaseTypes.EXLangSAccess __access = Microsoft.XLANGs.BaseTypes.EXLangSAccess.ePublic;
        #region port reflection support
        static public Microsoft.XLANGs.Core.OperationInfo SendToMsgBox = new Microsoft.XLANGs.Core.OperationInfo
        (
            "SendToMsgBox",
            System.Web.Services.Description.OperationFlow.RequestResponse,
            typeof(SendRequest_PortType),
            typeof(__messagetype_System_Xml_XmlDocument),
            typeof(__messagetype_System_Xml_XmlDocument),
            new System.Type[]{},
            new string[]{}
        );
        static public System.Collections.Hashtable OperationsInformation
        {
            get
            {
                System.Collections.Hashtable h = new System.Collections.Hashtable();
                h[ "SendToMsgBox" ] = SendToMsgBox;
                return h;
            }
        }
        #endregion // port reflection support
    }
    [Microsoft.XLANGs.BaseTypes.CorrelationTypeAttribute(
        Microsoft.XLANGs.BaseTypes.EXLangSAccess.eInternal,
        new string[] {
            "MSC.IntegrationService.Schemas.PropertySchema.provider", 
            "MSC.IntegrationService.Schemas.PropertySchema.recipient", 
            "MSC.IntegrationService.Schemas.PropertySchema.requestId", 
            "MSC.IntegrationService.Schemas.PropertySchema.requestSource", 
            "MSC.IntegrationService.Schemas.PropertySchema.serviceType"
        }
    )]
    sealed internal class PromotedProp_CorrelationType : Microsoft.XLANGs.Core.CorrelationType
    {
        public static readonly Microsoft.XLANGs.BaseTypes.EXLangSAccess __access = Microsoft.XLANGs.BaseTypes.EXLangSAccess.eInternal;
        private static Microsoft.XLANGs.BaseTypes.PropertyBase[] _properties = new Microsoft.XLANGs.BaseTypes.PropertyBase[] {
            new MSC.IntegrationService.Schemas.PropertySchema.provider(), 
            new MSC.IntegrationService.Schemas.PropertySchema.recipient(), 
            new MSC.IntegrationService.Schemas.PropertySchema.requestId(), 
            new MSC.IntegrationService.Schemas.PropertySchema.requestSource(), 
            new MSC.IntegrationService.Schemas.PropertySchema.serviceType()
         };
        public override Microsoft.XLANGs.BaseTypes.PropertyBase[] Properties { get { return _properties; } }
        public static Microsoft.XLANGs.BaseTypes.PropertyBase[] PropertiesList { get { return _properties; } }
    }
    //#line 398 "C:\MSC Products\Integration Service\Code\MSC.IntegrationService\MSC.IntegrationService\Orchestration.odx"
    [Microsoft.XLANGs.BaseTypes.StaticSubscriptionAttribute(
        0, "MSCReceive_Port", "ProcessRequest", -1, -1, true
    )]
    [Microsoft.XLANGs.BaseTypes.ServicePortsAttribute(
        new Microsoft.XLANGs.BaseTypes.EXLangSParameter[] {
            Microsoft.XLANGs.BaseTypes.EXLangSParameter.ePort|Microsoft.XLANGs.BaseTypes.EXLangSParameter.eImplements,
            Microsoft.XLANGs.BaseTypes.EXLangSParameter.ePort|Microsoft.XLANGs.BaseTypes.EXLangSParameter.eUses
        },
        new System.Type[] {
            typeof(MSC.IntegrationService.MSCReceive_PortType),
            typeof(MSC.IntegrationService.SendRequest_PortType)
        },
        new System.String[] {
            "MSCReceive_Port",
            "SendRequest_Port"
        },
        new System.Type[] {
            null,
            null
        }
    )]
    [Microsoft.XLANGs.BaseTypes.ServiceCallTreeAttribute(
        new System.Type[] {
        },
        new System.Type[] {
        },
        new System.Type[] {
        }
    )]
    [Microsoft.XLANGs.BaseTypes.ServiceAttribute(
        Microsoft.XLANGs.BaseTypes.EXLangSAccess.eInternal,
        Microsoft.XLANGs.BaseTypes.EXLangSServiceInfo.eNone
    )]
    [System.SerializableAttribute]
    [Microsoft.XLANGs.BaseTypes.BPELExportableAttribute(false)]
    sealed internal class MSCISOrchestration : Microsoft.BizTalk.XLANGs.BTXEngine.BTXService
    {
        public static readonly Microsoft.XLANGs.BaseTypes.EXLangSAccess __access = Microsoft.XLANGs.BaseTypes.EXLangSAccess.eInternal;
        public static readonly bool __execable = false;
        [Microsoft.XLANGs.BaseTypes.CallCompensationAttribute(
            Microsoft.XLANGs.BaseTypes.EXLangSCallCompensationInfo.eHasRequestResponse
,
            new System.String[] {
            },
            new System.String[] {
            }
        )]
        public static void __bodyProxy()
        {
        }
        private static System.Guid _serviceId = Microsoft.XLANGs.Core.HashHelper.HashServiceType(typeof(MSCISOrchestration));
        private static volatile System.Guid[] _activationSubIds;

        private static new object _lockIdentity = new object();

        public static System.Guid UUID { get { return _serviceId; } }
        public override System.Guid ServiceId { get { return UUID; } }

        protected override System.Guid[] ActivationSubGuids
        {
            get { return _activationSubIds; }
            set { _activationSubIds = value; }
        }

        protected override object StaleStateLock
        {
            get { return _lockIdentity; }
        }

        protected override bool HasActivation { get { return true; } }

        internal bool IsExeced = false;

        static MSCISOrchestration()
        {
            Microsoft.BizTalk.XLANGs.BTXEngine.BTXService.CacheStaticState( _serviceId );
        }

        private void ConstructorHelper()
        {
            _segments = new Microsoft.XLANGs.Core.Segment[] {
                new Microsoft.XLANGs.Core.Segment( new Microsoft.XLANGs.Core.Segment.SegmentCode(this.segment0), 0, 0, 0),
                new Microsoft.XLANGs.Core.Segment( new Microsoft.XLANGs.Core.Segment.SegmentCode(this.segment1), 1, 1, 1),
                new Microsoft.XLANGs.Core.Segment( new Microsoft.XLANGs.Core.Segment.SegmentCode(this.segment2), 1, 2, 2),
                new Microsoft.XLANGs.Core.Segment( new Microsoft.XLANGs.Core.Segment.SegmentCode(this.segment3), 1, 2, 3)
            };

            _Locks = 0;
            _rootContext = new __MSCISOrchestration_root_0(this);
            _stateMgrs = new Microsoft.XLANGs.Core.IStateManager[3];
            _stateMgrs[0] = _rootContext;
            FinalConstruct();
        }

        public MSCISOrchestration(System.Guid instanceId, Microsoft.BizTalk.XLANGs.BTXEngine.BTXSession session, Microsoft.BizTalk.XLANGs.BTXEngine.BTXEvents tracker)
            : base(instanceId, session, "MSCISOrchestration", tracker)
        {
            ConstructorHelper();
        }

        public MSCISOrchestration(int callIndex, System.Guid instanceId, Microsoft.BizTalk.XLANGs.BTXEngine.BTXService parent)
            : base(callIndex, instanceId, parent, "MSCISOrchestration")
        {
            ConstructorHelper();
        }

        private const string _symInfo = @"
<XsymFile>
<ProcessFlow xmlns:om='http://schemas.microsoft.com/BizTalk/2003/DesignerData'>      <shapeType>RootShape</shapeType>      <ShapeID>d6e10124-038a-49b5-ac5f-3ce96388591a</ShapeID>      
<children>                          
<ShapeInfo>      <shapeType>ReceiveShape</shapeType>      <ShapeID>15df91f9-3b33-446e-9515-c95751f7ce28</ShapeID>      <ParentLink>ServiceBody_Statement</ParentLink>                <shapeText>Receive_MSCMessage</shapeText>                  
<children>                </children>
  </ShapeInfo>
                            
<ShapeInfo>      <shapeType>ScopeShape</shapeType>      <ShapeID>5e3ab90f-5490-4bde-a650-b3ffe1d233d6</ShapeID>      <ParentLink>ServiceBody_Statement</ParentLink>                <shapeText>OverallOrchScope</shapeText>                  
<children>                          
<ShapeInfo>      <shapeType>VariableAssignmentShape</shapeType>      <ShapeID>4f10caf6-44a4-4b5b-86ee-982deebe1512</ShapeID>      <ParentLink>ComplexStatement_Statement</ParentLink>                <shapeText>Assign_Variables</shapeText>                  
<children>                </children>
  </ShapeInfo>
                            
<ShapeInfo>      <shapeType>VariableAssignmentShape</shapeType>      <ShapeID>12016790-ab78-4f33-8d6c-de986d56bd02</ShapeID>      <ParentLink>ComplexStatement_Statement</ParentLink>                <shapeText>Construct_RequestMsg</shapeText>                  
<children>                </children>
  </ShapeInfo>
                            
<ShapeInfo>      <shapeType>SendShape</shapeType>      <ShapeID>654358dc-de7d-40df-af29-d05600515104</ShapeID>      <ParentLink>ComplexStatement_Statement</ParentLink>                <shapeText>Send_RequestMsg</shapeText>                  
<children>                </children>
  </ShapeInfo>
                            
<ShapeInfo>      <shapeType>ReceiveShape</shapeType>      <ShapeID>7b385e9c-91c9-463f-8ad2-ed01f3c8cfb1</ShapeID>      <ParentLink>ComplexStatement_Statement</ParentLink>                <shapeText>Receive_ResponseMsg</shapeText>                  
<children>                </children>
  </ShapeInfo>
                            
<ShapeInfo>      <shapeType>VariableAssignmentShape</shapeType>      <ShapeID>c3a10cd2-1ede-4ddd-9e67-ece9a1083e03</ShapeID>      <ParentLink>ComplexStatement_Statement</ParentLink>                <shapeText>Construct_ResponseMsg</shapeText>                  
<children>                </children>
  </ShapeInfo>
                            
<ShapeInfo>      <shapeType>CatchShape</shapeType>      <ShapeID>4425b88f-7054-41ca-be97-5164d5048958</ShapeID>      <ParentLink>Scope_Catch</ParentLink>                <shapeText>OrchestrationException</shapeText>                      <ExceptionType>System.Exception</ExceptionType>            
<children>                          
<ShapeInfo>      <shapeType>ConstructShape</shapeType>      <ShapeID>a780fd31-1a7b-43bc-bed6-76bbfc237189</ShapeID>      <ParentLink>Catch_Statement</ParentLink>                <shapeText>Construct_ErrorMsg</shapeText>                  
<children>                          
<ShapeInfo>      <shapeType>MessageAssignmentShape</shapeType>      <ShapeID>c62c90af-1de5-40f1-9fb8-9bdb266cef14</ShapeID>      <ParentLink>ComplexStatement_Statement</ParentLink>                <shapeText>MessageAssignment_ErrMsg</shapeText>                  
<children>                </children>
  </ShapeInfo>
                            
<ShapeInfo>      <shapeType>MessageRefShape</shapeType>      <ShapeID>e2c3ec30-acbd-448f-b0d8-6a768abe3dd6</ShapeID>      <ParentLink>Construct_MessageRef</ParentLink>                  
<children>                </children>
  </ShapeInfo>
                  </children>
  </ShapeInfo>
                  </children>
  </ShapeInfo>
                  </children>
  </ShapeInfo>
                            
<ShapeInfo>      <shapeType>SendShape</shapeType>      <ShapeID>0151f644-04db-442b-a38c-8b3fa502ae93</ShapeID>      <ParentLink>ServiceBody_Statement</ParentLink>                <shapeText>Send_MSCResponseMsg</shapeText>                  
<children>                </children>
  </ShapeInfo>
                  </children>
  </ProcessFlow>
<Metadata>

<TrkMetadata>
<ActionName>'MSCISOrchestration'</ActionName><IsAtomic>0</IsAtomic><Line>398</Line><Position>14</Position><ShapeID>'e211a116-cb8b-44e7-a052-0de295aa0001'</ShapeID>
</TrkMetadata>

<TrkMetadata>
<Line>425</Line><Position>22</Position><ShapeID>'15df91f9-3b33-446e-9515-c95751f7ce28'</ShapeID>
<Messages>
	<MsgInfo><name>MSCMessage</name><part>part</part><schema>MSC.IntegrationService.Schemas.cdiInterop</schema><direction>Out</direction></MsgInfo>
</Messages>
</TrkMetadata>

<TrkMetadata>
<ActionName>'??__scope33'</ActionName><IsAtomic>0</IsAtomic><Line>438</Line><Position>13</Position><ShapeID>'5e3ab90f-5490-4bde-a650-b3ffe1d233d6'</ShapeID>
<Messages>
</Messages>
</TrkMetadata>

<TrkMetadata>
<Line>443</Line><Position>59</Position><ShapeID>'4f10caf6-44a4-4b5b-86ee-982deebe1512'</ShapeID>
<Messages>
</Messages>
</TrkMetadata>

<TrkMetadata>
<Line>458</Line><Position>59</Position><ShapeID>'12016790-ab78-4f33-8d6c-de986d56bd02'</ShapeID>
<Messages>
</Messages>
</TrkMetadata>

<TrkMetadata>
<Line>533</Line><Position>21</Position><ShapeID>'654358dc-de7d-40df-af29-d05600515104'</ShapeID>
<Messages>
</Messages>
</TrkMetadata>

<TrkMetadata>
<Line>535</Line><Position>21</Position><ShapeID>'7b385e9c-91c9-463f-8ad2-ed01f3c8cfb1'</ShapeID>
<Messages>
</Messages>
</TrkMetadata>

<TrkMetadata>
<Line>537</Line><Position>21</Position><ShapeID>'c3a10cd2-1ede-4ddd-9e67-ece9a1083e03'</ShapeID>
<Messages>
</Messages>
</TrkMetadata>

<TrkMetadata>
<Line>558</Line><Position>21</Position><ShapeID>'4425b88f-7054-41ca-be97-5164d5048958'</ShapeID>
<Messages>
</Messages>
</TrkMetadata>

<TrkMetadata>
<Line>561</Line><Position>25</Position><ShapeID>'a780fd31-1a7b-43bc-bed6-76bbfc237189'</ShapeID>
<Messages>
</Messages>
</TrkMetadata>

<TrkMetadata>
<Line>573</Line><Position>13</Position><ShapeID>'0151f644-04db-442b-a38c-8b3fa502ae93'</ShapeID>
<Messages>
</Messages>
</TrkMetadata>
</Metadata>
</XsymFile>";

        public override string odXml { get { return _symODXML; } }

        private const string _symODXML = @"
<?xml version='1.0' encoding='utf-8' standalone='yes'?>
<om:MetaModel MajorVersion='1' MinorVersion='3' Core='2b131234-7959-458d-834f-2dc0769ce683' ScheduleModel='66366196-361d-448d-976f-cab5e87496d2' xmlns:om='http://schemas.microsoft.com/BizTalk/2003/DesignerData'>
    <om:Element Type='Module' OID='ab005a1d-df5d-4ad0-9fb7-3c2a5499e06d' LowerBound='1.1' HigherBound='202.1'>
        <om:Property Name='ReportToAnalyst' Value='True' />
        <om:Property Name='Name' Value='MSC.IntegrationService' />
        <om:Property Name='Signal' Value='False' />
        <om:Element Type='PortType' OID='2358f104-2782-4d2b-915b-cc4320e32837' ParentLink='Module_PortType' LowerBound='4.1' HigherBound='11.1'>
            <om:Property Name='Synchronous' Value='True' />
            <om:Property Name='TypeModifier' Value='Public' />
            <om:Property Name='ReportToAnalyst' Value='True' />
            <om:Property Name='Name' Value='MSCReceive_PortType' />
            <om:Property Name='Signal' Value='True' />
            <om:Element Type='OperationDeclaration' OID='c89c4c82-b6e8-4bcb-ba83-440ed5fdd049' ParentLink='PortType_OperationDeclaration' LowerBound='6.1' HigherBound='10.1'>
                <om:Property Name='OperationType' Value='RequestResponse' />
                <om:Property Name='ReportToAnalyst' Value='True' />
                <om:Property Name='Name' Value='ProcessRequest' />
                <om:Property Name='Signal' Value='True' />
                <om:Element Type='MessageRef' OID='3c36acd1-09b7-4432-9d42-3b378ec43126' ParentLink='OperationDeclaration_RequestMessageRef' LowerBound='8.13' HigherBound='8.31'>
                    <om:Property Name='Ref' Value='MSC.IntegrationService.Schemas.cdiInterop' />
                    <om:Property Name='ReportToAnalyst' Value='True' />
                    <om:Property Name='Name' Value='Request' />
                    <om:Property Name='Signal' Value='True' />
                </om:Element>
                <om:Element Type='MessageRef' OID='4fc6729a-d589-4a0f-a590-4ee8894c622a' ParentLink='OperationDeclaration_ResponseMessageRef' LowerBound='8.33' HigherBound='8.55'>
                    <om:Property Name='Ref' Value='System.Xml.XmlDocument' />
                    <om:Property Name='ReportToAnalyst' Value='True' />
                    <om:Property Name='Name' Value='Response' />
                    <om:Property Name='Signal' Value='True' />
                </om:Element>
            </om:Element>
        </om:Element>
        <om:Element Type='PortType' OID='5451e1f5-138b-402b-b7a5-4c9f00e6779f' ParentLink='Module_PortType' LowerBound='11.1' HigherBound='18.1'>
            <om:Property Name='Synchronous' Value='True' />
            <om:Property Name='TypeModifier' Value='Public' />
            <om:Property Name='ReportToAnalyst' Value='True' />
            <om:Property Name='Name' Value='SendRequest_PortType' />
            <om:Property Name='Signal' Value='False' />
            <om:Element Type='OperationDeclaration' OID='c29d34a4-a0de-49ac-a48b-f84a396526f0' ParentLink='PortType_OperationDeclaration' LowerBound='13.1' HigherBound='17.1'>
                <om:Property Name='OperationType' Value='RequestResponse' />
                <om:Property Name='ReportToAnalyst' Value='True' />
                <om:Property Name='Name' Value='SendToMsgBox' />
                <om:Property Name='Signal' Value='True' />
                <om:Element Type='MessageRef' OID='9e034556-d85c-4947-b2fb-eec29af7931d' ParentLink='OperationDeclaration_RequestMessageRef' LowerBound='15.13' HigherBound='15.35'>
                    <om:Property Name='Ref' Value='System.Xml.XmlDocument' />
                    <om:Property Name='ReportToAnalyst' Value='True' />
                    <om:Property Name='Name' Value='Request' />
                    <om:Property Name='Signal' Value='True' />
                </om:Element>
                <om:Element Type='MessageRef' OID='eea39b04-8c67-400b-b88c-8a6436f30809' ParentLink='OperationDeclaration_ResponseMessageRef' LowerBound='15.37' HigherBound='15.59'>
                    <om:Property Name='Ref' Value='System.Xml.XmlDocument' />
                    <om:Property Name='ReportToAnalyst' Value='True' />
                    <om:Property Name='Name' Value='Response' />
                    <om:Property Name='Signal' Value='False' />
                </om:Element>
            </om:Element>
        </om:Element>
        <om:Element Type='ServiceDeclaration' OID='d73167d0-c4d7-4e9d-86bb-c2407c320d6f' ParentLink='Module_ServiceDeclaration' LowerBound='22.1' HigherBound='201.1'>
            <om:Property Name='InitializedTransactionType' Value='False' />
            <om:Property Name='IsInvokable' Value='False' />
            <om:Property Name='TypeModifier' Value='Internal' />
            <om:Property Name='ReportToAnalyst' Value='True' />
            <om:Property Name='Name' Value='MSCISOrchestration' />
            <om:Property Name='Signal' Value='False' />
            <om:Element Type='VariableDeclaration' OID='ca198d42-4866-4e5b-9168-c9758d2cc61b' ParentLink='ServiceDeclaration_VariableDeclaration' LowerBound='34.1' HigherBound='35.1'>
                <om:Property Name='UseDefaultConstructor' Value='False' />
                <om:Property Name='Type' Value='System.String' />
                <om:Property Name='ParamDirection' Value='In' />
                <om:Property Name='ReportToAnalyst' Value='True' />
                <om:Property Name='Name' Value='requestId' />
                <om:Property Name='Signal' Value='True' />
            </om:Element>
            <om:Element Type='VariableDeclaration' OID='53c42f97-baaa-4503-b3a6-afcdf1c8d7e5' ParentLink='ServiceDeclaration_VariableDeclaration' LowerBound='35.1' HigherBound='36.1'>
                <om:Property Name='UseDefaultConstructor' Value='False' />
                <om:Property Name='Type' Value='System.String' />
                <om:Property Name='ParamDirection' Value='In' />
                <om:Property Name='ReportToAnalyst' Value='True' />
                <om:Property Name='Name' Value='serviceType' />
                <om:Property Name='Signal' Value='True' />
            </om:Element>
            <om:Element Type='VariableDeclaration' OID='96ce8a35-e99a-4ef1-a278-ecfc6e193046' ParentLink='ServiceDeclaration_VariableDeclaration' LowerBound='36.1' HigherBound='37.1'>
                <om:Property Name='UseDefaultConstructor' Value='False' />
                <om:Property Name='Type' Value='System.String' />
                <om:Property Name='ParamDirection' Value='In' />
                <om:Property Name='ReportToAnalyst' Value='True' />
                <om:Property Name='Name' Value='provider' />
                <om:Property Name='Signal' Value='True' />
            </om:Element>
            <om:Element Type='VariableDeclaration' OID='465d0721-77d6-46f3-90fd-7e1ecaf36ef9' ParentLink='ServiceDeclaration_VariableDeclaration' LowerBound='37.1' HigherBound='38.1'>
                <om:Property Name='UseDefaultConstructor' Value='False' />
                <om:Property Name='Type' Value='System.String' />
                <om:Property Name='ParamDirection' Value='In' />
                <om:Property Name='ReportToAnalyst' Value='True' />
                <om:Property Name='Name' Value='requestSource' />
                <om:Property Name='Signal' Value='True' />
            </om:Element>
            <om:Element Type='VariableDeclaration' OID='da541b1e-7ad5-4c90-887f-569d0d4f46ea' ParentLink='ServiceDeclaration_VariableDeclaration' LowerBound='38.1' HigherBound='39.1'>
                <om:Property Name='UseDefaultConstructor' Value='True' />
                <om:Property Name='Type' Value='MSC.IntegrationService.ConfigService.ConfigManager' />
                <om:Property Name='ParamDirection' Value='In' />
                <om:Property Name='ReportToAnalyst' Value='True' />
                <om:Property Name='Name' Value='configMgr' />
                <om:Property Name='Signal' Value='True' />
            </om:Element>
            <om:Element Type='VariableDeclaration' OID='270fcd29-166c-4deb-8f60-e353d8561403' ParentLink='ServiceDeclaration_VariableDeclaration' LowerBound='39.1' HigherBound='40.1'>
                <om:Property Name='UseDefaultConstructor' Value='True' />
                <om:Property Name='Type' Value='MSC.IntegrationService.ConfigService.MessageConfig' />
                <om:Property Name='ParamDirection' Value='In' />
                <om:Property Name='ReportToAnalyst' Value='True' />
                <om:Property Name='Name' Value='messageConfig' />
                <om:Property Name='Signal' Value='True' />
            </om:Element>
            <om:Element Type='VariableDeclaration' OID='5045893e-230d-42ff-b0dc-6b8ed3ea217c' ParentLink='ServiceDeclaration_VariableDeclaration' LowerBound='40.1' HigherBound='41.1'>
                <om:Property Name='UseDefaultConstructor' Value='False' />
                <om:Property Name='Type' Value='System.String' />
                <om:Property Name='ParamDirection' Value='In' />
                <om:Property Name='ReportToAnalyst' Value='True' />
                <om:Property Name='Name' Value='requestMapName' />
                <om:Property Name='Signal' Value='True' />
            </om:Element>
            <om:Element Type='VariableDeclaration' OID='1a1db885-4c63-41b4-978f-7881a32a0f24' ParentLink='ServiceDeclaration_VariableDeclaration' LowerBound='41.1' HigherBound='42.1'>
                <om:Property Name='UseDefaultConstructor' Value='False' />
                <om:Property Name='Type' Value='System.String' />
                <om:Property Name='ParamDirection' Value='In' />
                <om:Property Name='ReportToAnalyst' Value='True' />
                <om:Property Name='Name' Value='responseMapName' />
                <om:Property Name='Signal' Value='True' />
            </om:Element>
            <om:Element Type='VariableDeclaration' OID='d8ab50e5-4229-4637-a5c6-3c5a672415cd' ParentLink='ServiceDeclaration_VariableDeclaration' LowerBound='42.1' HigherBound='43.1'>
                <om:Property Name='UseDefaultConstructor' Value='False' />
                <om:Property Name='Type' Value='System.Type' />
                <om:Property Name='ParamDirection' Value='In' />
                <om:Property Name='ReportToAnalyst' Value='True' />
                <om:Property Name='Name' Value='requestMapType' />
                <om:Property Name='Signal' Value='True' />
            </om:Element>
            <om:Element Type='VariableDeclaration' OID='d7b1fb23-b4a8-44a0-9273-807774f8862c' ParentLink='ServiceDeclaration_VariableDeclaration' LowerBound='43.1' HigherBound='44.1'>
                <om:Property Name='UseDefaultConstructor' Value='False' />
                <om:Property Name='Type' Value='System.Type' />
                <om:Property Name='ParamDirection' Value='In' />
                <om:Property Name='ReportToAnalyst' Value='True' />
                <om:Property Name='Name' Value='responseMapType' />
                <om:Property Name='Signal' Value='True' />
            </om:Element>
            <om:Element Type='VariableDeclaration' OID='d803660e-e3da-4a23-8852-3050b0fc817c' ParentLink='ServiceDeclaration_VariableDeclaration' LowerBound='44.1' HigherBound='45.1'>
                <om:Property Name='UseDefaultConstructor' Value='True' />
                <om:Property Name='Type' Value='MSC.IntegrationService.ConfigService.ServiceType' />
                <om:Property Name='ParamDirection' Value='In' />
                <om:Property Name='ReportToAnalyst' Value='True' />
                <om:Property Name='Name' Value='ServiceType' />
                <om:Property Name='Signal' Value='True' />
            </om:Element>
            <om:Element Type='VariableDeclaration' OID='639d1020-37ed-42a8-b44e-7aafb9a897a7' ParentLink='ServiceDeclaration_VariableDeclaration' LowerBound='45.1' HigherBound='46.1'>
                <om:Property Name='UseDefaultConstructor' Value='True' />
                <om:Property Name='Type' Value='MSC.IntegrationService.ConfigService.Recipient' />
                <om:Property Name='ParamDirection' Value='In' />
                <om:Property Name='ReportToAnalyst' Value='True' />
                <om:Property Name='Name' Value='Recipient' />
                <om:Property Name='Signal' Value='True' />
            </om:Element>
            <om:Element Type='VariableDeclaration' OID='0a7af2e8-e08d-4173-8f74-1f2819ed9cb1' ParentLink='ServiceDeclaration_VariableDeclaration' LowerBound='46.1' HigherBound='47.1'>
                <om:Property Name='UseDefaultConstructor' Value='False' />
                <om:Property Name='Type' Value='System.String' />
                <om:Property Name='ParamDirection' Value='In' />
                <om:Property Name='ReportToAnalyst' Value='True' />
                <om:Property Name='Name' Value='recipient' />
                <om:Property Name='Signal' Value='True' />
            </om:Element>
            <om:Element Type='CorrelationDeclaration' OID='345bfae5-0ba0-4483-9336-a07baad37efe' ParentLink='ServiceDeclaration_CorrelationDeclaration' LowerBound='29.1' HigherBound='30.1'>
                <om:Property Name='Type' Value='MSC.IntegrationService.PromotedProp_CorrelationType' />
                <om:Property Name='ParamDirection' Value='In' />
                <om:Property Name='ReportToAnalyst' Value='True' />
                <om:Property Name='AnalystComments' Value='Correlation Set to Promote Properties' />
                <om:Property Name='Name' Value='PromotedProp_Correlation' />
                <om:Property Name='Signal' Value='True' />
                <om:Element Type='StatementRef' OID='d47ab864-eed9-4fc5-8ee1-c9703da6f6f8' ParentLink='CorrelationDeclaration_StatementRef' LowerBound='158.74' HigherBound='158.109'>
                    <om:Property Name='Initializes' Value='True' />
                    <om:Property Name='Ref' Value='654358dc-de7d-40df-af29-d05600515104' />
                    <om:Property Name='ReportToAnalyst' Value='True' />
                    <om:Property Name='Name' Value='StatementRef_1' />
                    <om:Property Name='Signal' Value='False' />
                </om:Element>
            </om:Element>
            <om:Element Type='MessageDeclaration' OID='f4c7cef3-8a05-4a69-a2b2-c4e22325a24b' ParentLink='ServiceDeclaration_MessageDeclaration' LowerBound='30.1' HigherBound='31.1'>
                <om:Property Name='Type' Value='MSC.IntegrationService.Schemas.cdiInterop' />
                <om:Property Name='ParamDirection' Value='In' />
                <om:Property Name='ReportToAnalyst' Value='True' />
                <om:Property Name='Name' Value='MSCMessage' />
                <om:Property Name='Signal' Value='True' />
            </om:Element>
            <om:Element Type='MessageDeclaration' OID='04f7205d-4754-4b5f-9230-94892dac685c' ParentLink='ServiceDeclaration_MessageDeclaration' LowerBound='31.1' HigherBound='32.1'>
                <om:Property Name='Type' Value='System.Xml.XmlDocument' />
                <om:Property Name='ParamDirection' Value='In' />
                <om:Property Name='ReportToAnalyst' Value='True' />
                <om:Property Name='Name' Value='requestMessage' />
                <om:Property Name='Signal' Value='True' />
            </om:Element>
            <om:Element Type='MessageDeclaration' OID='b79d15d0-03df-4d99-b6c0-108214f3b35b' ParentLink='ServiceDeclaration_MessageDeclaration' LowerBound='32.1' HigherBound='33.1'>
                <om:Property Name='Type' Value='System.Xml.XmlDocument' />
                <om:Property Name='ParamDirection' Value='In' />
                <om:Property Name='ReportToAnalyst' Value='True' />
                <om:Property Name='Name' Value='responseMessage' />
                <om:Property Name='Signal' Value='True' />
            </om:Element>
            <om:Element Type='MessageDeclaration' OID='586da171-7e14-4616-a9bb-61357a4807d4' ParentLink='ServiceDeclaration_MessageDeclaration' LowerBound='33.1' HigherBound='34.1'>
                <om:Property Name='Type' Value='System.Xml.XmlDocument' />
                <om:Property Name='ParamDirection' Value='In' />
                <om:Property Name='ReportToAnalyst' Value='True' />
                <om:Property Name='Name' Value='MSCResponseMsg' />
                <om:Property Name='Signal' Value='True' />
            </om:Element>
            <om:Element Type='ServiceBody' OID='d6e10124-038a-49b5-ac5f-3ce96388591a' ParentLink='ServiceDeclaration_ServiceBody'>
                <om:Property Name='Signal' Value='False' />
                <om:Element Type='Receive' OID='15df91f9-3b33-446e-9515-c95751f7ce28' ParentLink='ServiceBody_Statement' LowerBound='49.1' HigherBound='62.1'>
                    <om:Property Name='Activate' Value='True' />
                    <om:Property Name='PortName' Value='MSCReceive_Port' />
                    <om:Property Name='MessageName' Value='MSCMessage' />
                    <om:Property Name='OperationName' Value='ProcessRequest' />
                    <om:Property Name='OperationMessageName' Value='Request' />
                    <om:Property Name='ReportToAnalyst' Value='True' />
                    <om:Property Name='Name' Value='Receive_MSCMessage' />
                    <om:Property Name='Signal' Value='True' />
                </om:Element>
                <om:Element Type='Scope' OID='5e3ab90f-5490-4bde-a650-b3ffe1d233d6' ParentLink='ServiceBody_Statement' LowerBound='62.1' HigherBound='197.1'>
                    <om:Property Name='InitializedTransactionType' Value='True' />
                    <om:Property Name='IsSynchronized' Value='False' />
                    <om:Property Name='ReportToAnalyst' Value='True' />
                    <om:Property Name='Name' Value='OverallOrchScope' />
                    <om:Property Name='Signal' Value='True' />
                    <om:Element Type='VariableAssignment' OID='4f10caf6-44a4-4b5b-86ee-982deebe1512' ParentLink='ComplexStatement_Statement' LowerBound='67.1' HigherBound='82.1'>
                        <om:Property Name='Expression' Value='System.Diagnostics.EventLog.WriteEntry(&quot;MSCIS&quot;, &quot;Received request, assigning variables&quot;);&#xD;&#xA;&#xD;&#xA;//get all promoted property values from incoming message&#xD;&#xA;requestId = MSCMessage(MSC.IntegrationService.Schemas.PropertySchema.requestId);&#xD;&#xA;serviceType = MSCMessage(MSC.IntegrationService.Schemas.PropertySchema.serviceType);&#xD;&#xA;requestSource = MSCMessage(MSC.IntegrationService.Schemas.PropertySchema.requestSource);&#xD;&#xA;provider = MSCMessage(MSC.IntegrationService.Schemas.PropertySchema.provider);&#xD;&#xA;recipient = MSCMessage(MSC.IntegrationService.Schemas.PropertySchema.recipient);&#xD;&#xA;&#xD;&#xA;//get message configuration based on provider of the request&#xD;&#xA;messageConfig = configMgr.GetMessageConfig(provider);&#xD;&#xA;ServiceType = messageConfig.ServiceTypes.GetServicType(serviceType);&#xD;&#xA;Recipient = ServiceType.RecipientList.GetRecipient(recipient);&#xD;&#xA;&#xD;&#xA;' />
                        <om:Property Name='ReportToAnalyst' Value='True' />
                        <om:Property Name='Name' Value='Assign_Variables' />
                        <om:Property Name='Signal' Value='False' />
                    </om:Element>
                    <om:Element Type='VariableAssignment' OID='12016790-ab78-4f33-8d6c-de986d56bd02' ParentLink='ComplexStatement_Statement' LowerBound='82.1' HigherBound='157.1'>
                        <om:Property Name='Expression' Value='System.Diagnostics.EventLog.WriteEntry(&quot;MSCIS&quot;, &quot;In Construct_RequestMsg&quot;);&#xD;&#xA;&#xD;&#xA;//get map name from message config. it should be an assembly type, for example&#xD;&#xA;//mapName = &quot;BTSSchemas.MSC_XYZ, BTSSchemas, Version=1.0.0.0, Culture=neutral, PublicKeyToken=19eedde3bd07a1d2&quot;;&#xD;&#xA;requestMapType = System.Type.GetType(Recipient.RequestMapType);&#xD;&#xA;&#xD;&#xA;//get the protocol and assign adapter specific properties while constructing message&#xD;&#xA;/*&#xD;&#xA;if (Recipient.RequestMapType.Length &gt; 0 )&#xD;&#xA;{&#xD;&#xA;    construct requestMessage&#xD;&#xA;    {&#xD;&#xA;        System.Diagnostics.EventLog.WriteEntry(&quot;MSCIS&quot;, &quot;Transforming.&quot;);&#xD;&#xA;        transform (requestMessage) = requestMapType(MSCMessage);&#xD;&#xA;    }&#xD;&#xA;}&#xD;&#xA;*/&#xD;&#xA;if (Recipient.Protocol == &quot;SOAP&quot;)&#xD;&#xA;{&#xD;&#xA;    construct requestMessage&#xD;&#xA;    {&#xD;&#xA;        System.Diagnostics.EventLog.WriteEntry(&quot;MSCIS&quot;, &quot;SOAP protocol, assigning parms.&quot;);&#xD;&#xA;        transform (requestMessage) = requestMapType(MSCMessage);&#xD;&#xA;        requestMessage(BTS.OutboundTransportType) = &quot;SOAP&quot;;&#xD;&#xA;        //make sure the following are correctly assigned - NC&#xD;&#xA;        requestMessage(SOAP.AssemblyName) = Recipient.Assembly ;&#xD;&#xA;        requestMessage(SOAP.MethodName) = Recipient.Operation;&#xD;&#xA;        requestMessage(SOAP.Username) = Recipient.UserId;&#xD;&#xA;        requestMessage(SOAP.Password) = Recipient.Password;&#xD;&#xA;        //requestMessage(BTS.Operation) = &quot;ExternalProcess&quot;;&#xD;&#xA;        //requestMessage(BTSSchemas.PropertySchema.ID) = ID;&#xD;&#xA;        &#xD;&#xA;    }&#xD;&#xA;    &#xD;&#xA;}&#xD;&#xA;else if (Recipient.Protocol == &quot;WSE&quot;)&#xD;&#xA;{&#xD;&#xA;    construct requestMessage&#xD;&#xA;    {&#xD;&#xA;        System.Diagnostics.EventLog.WriteEntry"+
@"(&quot;MSCIS&quot;, &quot;WSE protocol, assigning parms.&quot;);&#xD;&#xA;        transform (requestMessage) = requestMapType(MSCMessage);&#xD;&#xA;        requestMessage(BTS.OutboundTransportType) = &quot;WSE&quot;;&#xD;&#xA;        requestMessage(WSE.SoapAction) = &quot;&quot;;&#xD;&#xA;        requestMessage(WSE.ConfidentialityUser) = Recipient.UserId;&#xD;&#xA;        requestMessage(WSE.ConfidentialityPassword) = Recipient.Password;&#xD;&#xA;        requestMessage(WSE.IntegrityPasswordOption) = &quot;Hashed&quot;;&#xD;&#xA;        &#xD;&#xA;    }&#xD;&#xA;}&#xD;&#xA;else if (Recipient.Protocol == &quot;HTTP&quot;)&#xD;&#xA;{&#xD;&#xA;    construct requestMessage&#xD;&#xA;    {&#xD;&#xA;        System.Diagnostics.EventLog.WriteEntry(&quot;MSCIS&quot;, &quot;HTTP protocol, assigning parms.&quot;);&#xD;&#xA;        transform (requestMessage) = requestMapType(MSCMessage);&#xD;&#xA;           &#xD;&#xA;        requestMessage(BTS.OutboundTransportType) = &quot;HTTP&quot;;&#xD;&#xA;        requestMessage(MSC.IntegrationService.Schemas.PropertySchema.provider) = provider;&#xD;&#xA;        requestMessage(MSC.IntegrationService.Schemas.PropertySchema.recipient) = recipient;&#xD;&#xA;        requestMessage(MSC.IntegrationService.Schemas.PropertySchema.requestId) = requestId;&#xD;&#xA;        requestMessage(MSC.IntegrationService.Schemas.PropertySchema.requestSource) = requestSource;&#xD;&#xA;        requestMessage(MSC.IntegrationService.Schemas.PropertySchema.serviceType) = serviceType;&#xD;&#xA;    }&#xD;&#xA;}&#xD;&#xA;else&#xD;&#xA;{&#xD;&#xA;    construct requestMessage&#xD;&#xA;    {&#xD;&#xA;         System.Diagnostics.EventLog.WriteEntry(&quot;MSCIS&quot;, &quot;NO protocol.&quot;);&#xD;&#xA;         transform (requestMessage) = requestMapType(MSCMessage);&#xD;&#xA;         //default protocol&#xD;&#xA;    }&#xD;&#xA;}&#xD;&#xA;&#xD;&#xA;' />
                        <om:Property Name='ReportToAnalyst' Value='True' />
                        <om:Property Name='Name' Value='Construct_RequestMsg' />
                        <om:Property Name='Signal' Value='False' />
                    </om:Element>
                    <om:Element Type='Send' OID='654358dc-de7d-40df-af29-d05600515104' ParentLink='ComplexStatement_Statement' LowerBound='157.1' HigherBound='159.1'>
                        <om:Property Name='PortName' Value='SendRequest_Port' />
                        <om:Property Name='MessageName' Value='requestMessage' />
                        <om:Property Name='OperationName' Value='SendToMsgBox' />
                        <om:Property Name='OperationMessageName' Value='Request' />
                        <om:Property Name='ReportToAnalyst' Value='True' />
                        <om:Property Name='Name' Value='Send_RequestMsg' />
                        <om:Property Name='Signal' Value='True' />
                    </om:Element>
                    <om:Element Type='Receive' OID='7b385e9c-91c9-463f-8ad2-ed01f3c8cfb1' ParentLink='ComplexStatement_Statement' LowerBound='159.1' HigherBound='161.1'>
                        <om:Property Name='Activate' Value='False' />
                        <om:Property Name='PortName' Value='SendRequest_Port' />
                        <om:Property Name='MessageName' Value='responseMessage' />
                        <om:Property Name='OperationName' Value='SendToMsgBox' />
                        <om:Property Name='OperationMessageName' Value='Response' />
                        <om:Property Name='ReportToAnalyst' Value='True' />
                        <om:Property Name='Name' Value='Receive_ResponseMsg' />
                        <om:Property Name='Signal' Value='True' />
                    </om:Element>
                    <om:Element Type='VariableAssignment' OID='c3a10cd2-1ede-4ddd-9e67-ece9a1083e03' ParentLink='ComplexStatement_Statement' LowerBound='161.1' HigherBound='179.1'>
                        <om:Property Name='Expression' Value='if (Recipient.ResponseMapType != &quot;&quot;)&#xD;&#xA;{&#xD;&#xA;    construct MSCResponseMsg&#xD;&#xA;    {&#xD;&#xA;        //get responseMapType using message config&#xD;&#xA;        responseMapType = System.Type.GetType(Recipient.ResponseMapType);&#xD;&#xA;        transform (MSCResponseMsg) = responseMapType(responseMessage);&#xD;&#xA;        &#xD;&#xA;    }&#xD;&#xA;}&#xD;&#xA;else&#xD;&#xA;{&#xD;&#xA;    construct MSCResponseMsg&#xD;&#xA;        {     &#xD;&#xA;            MSCResponseMsg = responseMessage;&#xD;&#xA;        }&#xD;&#xA;}' />
                        <om:Property Name='ReportToAnalyst' Value='True' />
                        <om:Property Name='Name' Value='Construct_ResponseMsg' />
                        <om:Property Name='Signal' Value='True' />
                    </om:Element>
                    <om:Element Type='Catch' OID='4425b88f-7054-41ca-be97-5164d5048958' ParentLink='Scope_Catch' LowerBound='182.1' HigherBound='195.1'>
                        <om:Property Name='ExceptionName' Value='ex' />
                        <om:Property Name='ExceptionType' Value='System.Exception' />
                        <om:Property Name='IsFaultMessage' Value='False' />
                        <om:Property Name='ReportToAnalyst' Value='True' />
                        <om:Property Name='Name' Value='OrchestrationException' />
                        <om:Property Name='Signal' Value='True' />
                        <om:Element Type='Construct' OID='a780fd31-1a7b-43bc-bed6-76bbfc237189' ParentLink='Catch_Statement' LowerBound='185.1' HigherBound='194.1'>
                            <om:Property Name='ReportToAnalyst' Value='True' />
                            <om:Property Name='Name' Value='Construct_ErrorMsg' />
                            <om:Property Name='Signal' Value='True' />
                            <om:Element Type='MessageAssignment' OID='c62c90af-1de5-40f1-9fb8-9bdb266cef14' ParentLink='ComplexStatement_Statement' LowerBound='188.1' HigherBound='193.1'>
                                <om:Property Name='Expression' Value='System.Diagnostics.EventLog.WriteEntry(&quot;MSCIS&quot;, &quot;Error occurred:&quot; + ex.Message);&#xD;&#xA;&#xD;&#xA;MSCResponseMsg = MSCMessage;&#xD;&#xA;xpath(MSCResponseMsg, &quot;/*[local-name()=&apos;clipboard&apos;]/*[local-name()=&apos;msgRequest&apos;]/*[local-name()=&apos;msgErrors&apos;]/*[local-name()=&apos;msgError&apos;]/@*[local-name()=&apos;message&apos;]&quot;)= ex.Message;' />
                                <om:Property Name='ReportToAnalyst' Value='False' />
                                <om:Property Name='Name' Value='MessageAssignment_ErrMsg' />
                                <om:Property Name='Signal' Value='True' />
                            </om:Element>
                            <om:Element Type='MessageRef' OID='e2c3ec30-acbd-448f-b0d8-6a768abe3dd6' ParentLink='Construct_MessageRef' LowerBound='186.35' HigherBound='186.49'>
                                <om:Property Name='Ref' Value='MSCResponseMsg' />
                                <om:Property Name='ReportToAnalyst' Value='True' />
                                <om:Property Name='Signal' Value='False' />
                            </om:Element>
                        </om:Element>
                    </om:Element>
                </om:Element>
                <om:Element Type='Send' OID='0151f644-04db-442b-a38c-8b3fa502ae93' ParentLink='ServiceBody_Statement' LowerBound='197.1' HigherBound='199.1'>
                    <om:Property Name='PortName' Value='MSCReceive_Port' />
                    <om:Property Name='MessageName' Value='MSCResponseMsg' />
                    <om:Property Name='OperationName' Value='ProcessRequest' />
                    <om:Property Name='OperationMessageName' Value='Response' />
                    <om:Property Name='ReportToAnalyst' Value='True' />
                    <om:Property Name='Name' Value='Send_MSCResponseMsg' />
                    <om:Property Name='Signal' Value='True' />
                </om:Element>
            </om:Element>
            <om:Element Type='PortDeclaration' OID='9891102d-66ee-4cef-83d6-1ed77dd0de09' ParentLink='ServiceDeclaration_PortDeclaration' LowerBound='25.1' HigherBound='27.1'>
                <om:Property Name='PortModifier' Value='Implements' />
                <om:Property Name='Orientation' Value='Left' />
                <om:Property Name='PortIndex' Value='6' />
                <om:Property Name='IsWebPort' Value='False' />
                <om:Property Name='OrderedDelivery' Value='False' />
                <om:Property Name='DeliveryNotification' Value='None' />
                <om:Property Name='Type' Value='MSC.IntegrationService.MSCReceive_PortType' />
                <om:Property Name='ParamDirection' Value='In' />
                <om:Property Name='ReportToAnalyst' Value='True' />
                <om:Property Name='Name' Value='MSCReceive_Port' />
                <om:Property Name='Signal' Value='False' />
                <om:Element Type='LogicalBindingAttribute' OID='b599bfac-fd70-4b39-8216-3a44855929ae' ParentLink='PortDeclaration_CLRAttribute' LowerBound='25.1' HigherBound='26.1'>
                    <om:Property Name='Signal' Value='False' />
                </om:Element>
            </om:Element>
            <om:Element Type='PortDeclaration' OID='40c1823c-dc5a-4bb8-a44a-3e56d3b20902' ParentLink='ServiceDeclaration_PortDeclaration' LowerBound='27.1' HigherBound='29.1'>
                <om:Property Name='PortModifier' Value='Uses' />
                <om:Property Name='Orientation' Value='Right' />
                <om:Property Name='PortIndex' Value='17' />
                <om:Property Name='IsWebPort' Value='False' />
                <om:Property Name='OrderedDelivery' Value='False' />
                <om:Property Name='DeliveryNotification' Value='None' />
                <om:Property Name='Type' Value='MSC.IntegrationService.SendRequest_PortType' />
                <om:Property Name='ParamDirection' Value='In' />
                <om:Property Name='ReportToAnalyst' Value='True' />
                <om:Property Name='Name' Value='SendRequest_Port' />
                <om:Property Name='Signal' Value='False' />
                <om:Element Type='DirectBindingAttribute' OID='bf73a900-e126-4802-beb0-c5343e4ecb81' ParentLink='PortDeclaration_CLRAttribute' LowerBound='27.1' HigherBound='28.1'>
                    <om:Property Name='DirectBindingType' Value='MessageBox' />
                    <om:Property Name='Signal' Value='False' />
                </om:Element>
            </om:Element>
        </om:Element>
        <om:Element Type='CorrelationType' OID='9992c641-a022-45a2-8fb1-aa3af1473cca' ParentLink='Module_CorrelationType' LowerBound='18.1' HigherBound='22.1'>
            <om:Property Name='TypeModifier' Value='Internal' />
            <om:Property Name='ReportToAnalyst' Value='True' />
            <om:Property Name='Name' Value='PromotedProp_CorrelationType' />
            <om:Property Name='Signal' Value='True' />
            <om:Element Type='PropertyRef' OID='b5aa1521-be2d-49cc-bf9b-f9becce33c87' ParentLink='CorrelationType_PropertyRef' LowerBound='20.9' HigherBound='20.40'>
                <om:Property Name='Ref' Value='MSC.IntegrationService.Schemas.PropertySchema.provider' />
                <om:Property Name='ReportToAnalyst' Value='True' />
                <om:Property Name='Name' Value='PropertyRef_1' />
                <om:Property Name='Signal' Value='False' />
            </om:Element>
            <om:Element Type='PropertyRef' OID='09bee511-76ed-4288-97a6-213ffc61d504' ParentLink='CorrelationType_PropertyRef' LowerBound='20.42' HigherBound='20.74'>
                <om:Property Name='Ref' Value='MSC.IntegrationService.Schemas.PropertySchema.recipient' />
                <om:Property Name='ReportToAnalyst' Value='True' />
                <om:Property Name='Name' Value='PropertyRef_1' />
                <om:Property Name='Signal' Value='False' />
            </om:Element>
            <om:Element Type='PropertyRef' OID='68f83811-5610-4daa-aada-943a199583e6' ParentLink='CorrelationType_PropertyRef' LowerBound='20.76' HigherBound='20.108'>
                <om:Property Name='Ref' Value='MSC.IntegrationService.Schemas.PropertySchema.requestId' />
                <om:Property Name='ReportToAnalyst' Value='True' />
                <om:Property Name='Name' Value='PropertyRef_1' />
                <om:Property Name='Signal' Value='False' />
            </om:Element>
            <om:Element Type='PropertyRef' OID='c73dadb9-7cde-4eca-b68f-9e73684dcf9b' ParentLink='CorrelationType_PropertyRef' LowerBound='20.110' HigherBound='20.146'>
                <om:Property Name='Ref' Value='MSC.IntegrationService.Schemas.PropertySchema.requestSource' />
                <om:Property Name='ReportToAnalyst' Value='True' />
                <om:Property Name='Name' Value='PropertyRef_1' />
                <om:Property Name='Signal' Value='False' />
            </om:Element>
            <om:Element Type='PropertyRef' OID='b2f8692c-3261-49b9-801d-f5a72858266f' ParentLink='CorrelationType_PropertyRef' LowerBound='20.148' HigherBound='20.182'>
                <om:Property Name='Ref' Value='MSC.IntegrationService.Schemas.PropertySchema.serviceType' />
                <om:Property Name='ReportToAnalyst' Value='True' />
                <om:Property Name='Name' Value='PropertyRef_1' />
                <om:Property Name='Signal' Value='False' />
            </om:Element>
        </om:Element>
    </om:Element>
</om:MetaModel>
";

        [System.SerializableAttribute]
        public class __MSCISOrchestration_root_0 : Microsoft.XLANGs.Core.ServiceContext
        {
            public __MSCISOrchestration_root_0(Microsoft.XLANGs.Core.Service svc)
                : base(svc, "MSCISOrchestration")
            {
            }

            public override int Index { get { return 0; } }

            public override Microsoft.XLANGs.Core.Segment InitialSegment
            {
                get { return _service._segments[0]; }
            }
            public override Microsoft.XLANGs.Core.Segment FinalSegment
            {
                get { return _service._segments[0]; }
            }

            public override int CompensationSegment { get { return -1; } }
            public override bool OnError()
            {
                Finally();
                return false;
            }

            public override void Finally()
            {
                MSCISOrchestration __svc__ = (MSCISOrchestration)_service;
                __MSCISOrchestration_root_0 __ctx0__ = (__MSCISOrchestration_root_0)(__svc__._stateMgrs[0]);

                if (__svc__.MSCReceive_Port != null)
                {
                    __svc__.MSCReceive_Port.Close(this, null);
                    __svc__.MSCReceive_Port = null;
                }
                if (__svc__.SendRequest_Port != null)
                {
                    __svc__.SendRequest_Port.Close(this, null);
                    __svc__.SendRequest_Port = null;
                }
                base.Finally();
            }

            internal Microsoft.XLANGs.Core.SubscriptionWrapper __subWrapper0;
            internal Microsoft.XLANGs.Core.SubscriptionWrapper __subWrapper1;
        }


        [System.SerializableAttribute]
        public class __MSCISOrchestration_1 : Microsoft.XLANGs.Core.ExceptionHandlingContext
        {
            public __MSCISOrchestration_1(Microsoft.XLANGs.Core.Service svc)
                : base(svc, "MSCISOrchestration")
            {
            }

            public override int Index { get { return 1; } }

            public override bool CombineParentCommit { get { return true; } }

            public override Microsoft.XLANGs.Core.Segment InitialSegment
            {
                get { return _service._segments[1]; }
            }
            public override Microsoft.XLANGs.Core.Segment FinalSegment
            {
                get { return _service._segments[1]; }
            }

            public override int CompensationSegment { get { return -1; } }
            public override bool OnError()
            {
                Finally();
                return false;
            }

            public override void Finally()
            {
                MSCISOrchestration __svc__ = (MSCISOrchestration)_service;
                __MSCISOrchestration_1 __ctx1__ = (__MSCISOrchestration_1)(__svc__._stateMgrs[1]);

                if (__ctx1__ != null && __ctx1__.__MSCMessage != null)
                {
                    __ctx1__.UnrefMessage(__ctx1__.__MSCMessage);
                    __ctx1__.__MSCMessage = null;
                }
                if (__ctx1__ != null && __ctx1__.__MSCResponseMsg != null)
                {
                    __ctx1__.UnrefMessage(__ctx1__.__MSCResponseMsg);
                    __ctx1__.__MSCResponseMsg = null;
                }
                if (__ctx1__ != null)
                    __ctx1__.__requestId = null;
                if (__ctx1__ != null)
                    __ctx1__.__serviceType = null;
                if (__ctx1__ != null)
                    __ctx1__.__provider = null;
                if (__ctx1__ != null)
                    __ctx1__.__requestSource = null;
                if (__ctx1__ != null)
                    __ctx1__.__configMgr = null;
                if (__ctx1__ != null)
                    __ctx1__.__messageConfig = null;
                if (__ctx1__ != null)
                    __ctx1__.__requestMapName = null;
                if (__ctx1__ != null)
                    __ctx1__.__responseMapName = null;
                if (__ctx1__ != null)
                    __ctx1__.__ServiceType = null;
                if (__ctx1__ != null)
                    __ctx1__.__Recipient = null;
                if (__ctx1__ != null)
                    __ctx1__.__recipient = null;
                base.Finally();
            }

            [Microsoft.XLANGs.Core.UserVariableAttribute("MSCMessage")]
            public __messagetype_MSC_IntegrationService_Schemas_cdiInterop __MSCMessage;
            [Microsoft.XLANGs.Core.UserVariableAttribute("requestMessage")]
            public __messagetype_System_Xml_XmlDocument __requestMessage;
            [Microsoft.XLANGs.Core.UserVariableAttribute("responseMessage")]
            public __messagetype_System_Xml_XmlDocument __responseMessage;
            [Microsoft.XLANGs.Core.UserVariableAttribute("MSCResponseMsg")]
            public __messagetype_System_Xml_XmlDocument __MSCResponseMsg;
            [Microsoft.XLANGs.Core.UserVariableAttribute("PromotedProp_Correlation")]
            internal Microsoft.XLANGs.Core.Correlation __PromotedProp_Correlation;
            [Microsoft.XLANGs.Core.UserVariableAttribute("requestId")]
            internal System.String __requestId;
            [Microsoft.XLANGs.Core.UserVariableAttribute("serviceType")]
            internal System.String __serviceType;
            [Microsoft.XLANGs.Core.UserVariableAttribute("provider")]
            internal System.String __provider;
            [Microsoft.XLANGs.Core.UserVariableAttribute("requestSource")]
            internal System.String __requestSource;
            [Microsoft.XLANGs.Core.UserVariableAttribute("configMgr")]
            internal MSC.IntegrationService.ConfigService.ConfigManager __configMgr;
            [Microsoft.XLANGs.Core.UserVariableAttribute("messageConfig")]
            internal MSC.IntegrationService.ConfigService.MessageConfig __messageConfig;
            [Microsoft.XLANGs.Core.UserVariableAttribute("requestMapName")]
            internal System.String __requestMapName;
            [Microsoft.XLANGs.Core.UserVariableAttribute("responseMapName")]
            internal System.String __responseMapName;
            [Microsoft.XLANGs.Core.UserVariableAttribute("requestMapType")]
            internal System.Type __requestMapType;
            [Microsoft.XLANGs.Core.UserVariableAttribute("responseMapType")]
            internal System.Type __responseMapType;
            [Microsoft.XLANGs.Core.UserVariableAttribute("ServiceType")]
            internal MSC.IntegrationService.ConfigService.ServiceType __ServiceType;
            [Microsoft.XLANGs.Core.UserVariableAttribute("Recipient")]
            internal MSC.IntegrationService.ConfigService.Recipient __Recipient;
            [Microsoft.XLANGs.Core.UserVariableAttribute("recipient")]
            internal System.String __recipient;
        }


        [System.SerializableAttribute]
        public class ____scope33_2 : Microsoft.XLANGs.Core.ExceptionHandlingContext
        {
            public ____scope33_2(Microsoft.XLANGs.Core.Service svc)
                : base(svc, "??__scope33")
            {
            }

            public override int Index { get { return 2; } }

            public override Microsoft.XLANGs.Core.Segment InitialSegment
            {
                get { return _service._segments[2]; }
            }
            public override Microsoft.XLANGs.Core.Segment FinalSegment
            {
                get { return _service._segments[2]; }
            }

            public override int CompensationSegment { get { return -1; } }
            public override bool OnError()
            {
                Microsoft.XLANGs.Core.Segment __seg__;
                Microsoft.XLANGs.Core.FaultReceiveException __fault__;

                __exv__ = _exception;
                if (!(__exv__ is Microsoft.XLANGs.Core.UnknownException))
                {
                    __fault__ = __exv__ as Microsoft.XLANGs.Core.FaultReceiveException;
                    if ((__fault__ == null) && (__exv__ is System.Exception))
                    {
                        __seg__ = _service._segments[3];
                        __seg__.Reset(1);
                        __seg__.PredecessorDone(_service);
                        return true;
                    }
                }

                Finally();
                return false;
            }

            public override void Finally()
            {
                MSCISOrchestration __svc__ = (MSCISOrchestration)_service;
                __MSCISOrchestration_root_0 __ctx0__ = (__MSCISOrchestration_root_0)(__svc__._stateMgrs[0]);
                __MSCISOrchestration_1 __ctx1__ = (__MSCISOrchestration_1)(__svc__._stateMgrs[1]);
                ____scope33_2 __ctx2__ = (____scope33_2)(__svc__._stateMgrs[2]);

                if (__ctx1__ != null && __ctx1__.__requestMessage != null)
                {
                    __ctx1__.UnrefMessage(__ctx1__.__requestMessage);
                    __ctx1__.__requestMessage = null;
                }
                if (__ctx1__ != null && __ctx1__.__PromotedProp_Correlation != null)
                    __ctx1__.__PromotedProp_Correlation = null;
                if (__ctx1__ != null && __ctx1__.__responseMessage != null)
                {
                    __ctx1__.UnrefMessage(__ctx1__.__responseMessage);
                    __ctx1__.__responseMessage = null;
                }
                if (__ctx1__ != null)
                    __ctx1__.__requestMapType = null;
                if (__ctx1__ != null)
                    __ctx1__.__responseMapType = null;
                if (__ctx2__ != null)
                    __ctx2__.__ex_0 = null;
                if (__ctx0__ != null && __ctx0__.__subWrapper1 != null)
                {
                    __ctx0__.__subWrapper1.Destroy(__svc__, __ctx0__);
                    __ctx0__.__subWrapper1 = null;
                }
                base.Finally();
            }

            internal object __exv__;
            internal System.Exception __ex_0
            {
                get { return (System.Exception)__exv__; }
                set { __exv__ = value; }
            }
        }

        private static Microsoft.XLANGs.Core.CorrelationType[] _correlationTypes = new Microsoft.XLANGs.Core.CorrelationType[] { new PromotedProp_CorrelationType() };
        public override Microsoft.XLANGs.Core.CorrelationType[] CorrelationTypes { get { return _correlationTypes; } }

        private static System.Guid[] _convoySetIds;

        public override System.Guid[] ConvoySetGuids
        {
            get { return _convoySetIds; }
            set { _convoySetIds = value; }
        }

        public static object[] StaticConvoySetInformation
        {
            get {
                return null;
            }
        }

        [Microsoft.XLANGs.BaseTypes.LogicalBindingAttribute()]
        [Microsoft.XLANGs.BaseTypes.PortAttribute(
            Microsoft.XLANGs.BaseTypes.EXLangSParameter.eImplements
        )]
        [Microsoft.XLANGs.Core.UserVariableAttribute("MSCReceive_Port")]
        internal MSCReceive_PortType MSCReceive_Port;
        [Microsoft.XLANGs.BaseTypes.DirectBindingAttribute()]
        [Microsoft.XLANGs.BaseTypes.PortAttribute(
            Microsoft.XLANGs.BaseTypes.EXLangSParameter.eUses
        )]
        [Microsoft.XLANGs.Core.UserVariableAttribute("SendRequest_Port")]
        internal SendRequest_PortType SendRequest_Port;

        public static Microsoft.XLANGs.Core.PortInfo[] _portInfo = new Microsoft.XLANGs.Core.PortInfo[] {
            new Microsoft.XLANGs.Core.PortInfo(new Microsoft.XLANGs.Core.OperationInfo[] {MSCReceive_PortType.ProcessRequest},
                                               typeof(MSCISOrchestration).GetField("MSCReceive_Port", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance),
                                               Microsoft.XLANGs.BaseTypes.Polarity.implements,
                                               false,
                                               Microsoft.XLANGs.Core.HashHelper.HashPort(typeof(MSCISOrchestration), "MSCReceive_Port"),
                                               null),
            new Microsoft.XLANGs.Core.PortInfo(new Microsoft.XLANGs.Core.OperationInfo[] {SendRequest_PortType.SendToMsgBox},
                                               typeof(MSCISOrchestration).GetField("SendRequest_Port", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance),
                                               Microsoft.XLANGs.BaseTypes.Polarity.uses,
                                               false,
                                               Microsoft.XLANGs.Core.HashHelper.HashPort(typeof(MSCISOrchestration), "SendRequest_Port"),
                                               null)
        };

        public override Microsoft.XLANGs.Core.PortInfo[] PortInformation
        {
            get { return _portInfo; }
        }

        static public System.Collections.Hashtable PortsInformation
        {
            get
            {
                System.Collections.Hashtable h = new System.Collections.Hashtable();
                h[_portInfo[0].Name] = _portInfo[0];
                h[_portInfo[1].Name] = _portInfo[1];
                return h;
            }
        }

        public static System.Type[] InvokedServicesTypes
        {
            get
            {
                return new System.Type[] {
                    // type of each service invoked by this service
                };
            }
        }

        public static System.Type[] CalledServicesTypes
        {
            get
            {
                return new System.Type[] {
                };
            }
        }

        public static System.Type[] ExecedServicesTypes
        {
            get
            {
                return new System.Type[] {
                };
            }
        }

        public static object[] StaticSubscriptionsInformation {
            get {
                return new object[1]{
                     new object[5] { _portInfo[0], 0, null , -1, true }
                };
            }
        }

        public static Microsoft.XLANGs.RuntimeTypes.Location[] __eventLocations = new Microsoft.XLANGs.RuntimeTypes.Location[] {
            new Microsoft.XLANGs.RuntimeTypes.Location(0, "00000000-0000-0000-0000-000000000000", 1, true),
            new Microsoft.XLANGs.RuntimeTypes.Location(1, "15df91f9-3b33-446e-9515-c95751f7ce28", 1, true),
            new Microsoft.XLANGs.RuntimeTypes.Location(2, "15df91f9-3b33-446e-9515-c95751f7ce28", 1, false),
            new Microsoft.XLANGs.RuntimeTypes.Location(3, "00000000-0000-0000-0000-000000000000", 1, false),
            new Microsoft.XLANGs.RuntimeTypes.Location(4, "5e3ab90f-5490-4bde-a650-b3ffe1d233d6", 1, true),
            new Microsoft.XLANGs.RuntimeTypes.Location(5, "00000000-0000-0000-0000-000000000000", 2, true),
            new Microsoft.XLANGs.RuntimeTypes.Location(6, "4f10caf6-44a4-4b5b-86ee-982deebe1512", 2, true),
            new Microsoft.XLANGs.RuntimeTypes.Location(7, "4f10caf6-44a4-4b5b-86ee-982deebe1512", 2, false),
            new Microsoft.XLANGs.RuntimeTypes.Location(8, "00000000-0000-0000-0000-000000000000", 2, false),
            new Microsoft.XLANGs.RuntimeTypes.Location(9, "12016790-ab78-4f33-8d6c-de986d56bd02", 2, true),
            new Microsoft.XLANGs.RuntimeTypes.Location(10, "12016790-ab78-4f33-8d6c-de986d56bd02", 2, false),
            new Microsoft.XLANGs.RuntimeTypes.Location(11, "654358dc-de7d-40df-af29-d05600515104", 2, true),
            new Microsoft.XLANGs.RuntimeTypes.Location(12, "654358dc-de7d-40df-af29-d05600515104", 2, false),
            new Microsoft.XLANGs.RuntimeTypes.Location(13, "7b385e9c-91c9-463f-8ad2-ed01f3c8cfb1", 2, true),
            new Microsoft.XLANGs.RuntimeTypes.Location(14, "7b385e9c-91c9-463f-8ad2-ed01f3c8cfb1", 2, false),
            new Microsoft.XLANGs.RuntimeTypes.Location(15, "c3a10cd2-1ede-4ddd-9e67-ece9a1083e03", 2, true),
            new Microsoft.XLANGs.RuntimeTypes.Location(16, "c3a10cd2-1ede-4ddd-9e67-ece9a1083e03", 2, false),
            new Microsoft.XLANGs.RuntimeTypes.Location(17, "4425b88f-7054-41ca-be97-5164d5048958", 3, true),
            new Microsoft.XLANGs.RuntimeTypes.Location(18, "a780fd31-1a7b-43bc-bed6-76bbfc237189", 3, true),
            new Microsoft.XLANGs.RuntimeTypes.Location(19, "a780fd31-1a7b-43bc-bed6-76bbfc237189", 3, false),
            new Microsoft.XLANGs.RuntimeTypes.Location(20, "4425b88f-7054-41ca-be97-5164d5048958", 3, false),
            new Microsoft.XLANGs.RuntimeTypes.Location(21, "5e3ab90f-5490-4bde-a650-b3ffe1d233d6", 1, false),
            new Microsoft.XLANGs.RuntimeTypes.Location(22, "0151f644-04db-442b-a38c-8b3fa502ae93", 1, true),
            new Microsoft.XLANGs.RuntimeTypes.Location(23, "0151f644-04db-442b-a38c-8b3fa502ae93", 1, false)
        };

        public override Microsoft.XLANGs.RuntimeTypes.Location[] EventLocations
        {
            get { return __eventLocations; }
        }

        public static Microsoft.XLANGs.RuntimeTypes.EventData[] __eventData = new Microsoft.XLANGs.RuntimeTypes.EventData[] {
            new Microsoft.XLANGs.RuntimeTypes.EventData( Microsoft.XLANGs.RuntimeTypes.Operation.Start | Microsoft.XLANGs.RuntimeTypes.Operation.Body),
            new Microsoft.XLANGs.RuntimeTypes.EventData( Microsoft.XLANGs.RuntimeTypes.Operation.Start | Microsoft.XLANGs.RuntimeTypes.Operation.Receive),
            new Microsoft.XLANGs.RuntimeTypes.EventData( Microsoft.XLANGs.RuntimeTypes.Operation.Start | Microsoft.XLANGs.RuntimeTypes.Operation.Scope),
            new Microsoft.XLANGs.RuntimeTypes.EventData( Microsoft.XLANGs.RuntimeTypes.Operation.Start | Microsoft.XLANGs.RuntimeTypes.Operation.Expression),
            new Microsoft.XLANGs.RuntimeTypes.EventData( Microsoft.XLANGs.RuntimeTypes.Operation.End | Microsoft.XLANGs.RuntimeTypes.Operation.Expression),
            new Microsoft.XLANGs.RuntimeTypes.EventData( Microsoft.XLANGs.RuntimeTypes.Operation.Start | Microsoft.XLANGs.RuntimeTypes.Operation.If),
            new Microsoft.XLANGs.RuntimeTypes.EventData( Microsoft.XLANGs.RuntimeTypes.Operation.Start | Microsoft.XLANGs.RuntimeTypes.Operation.Construct),
            new Microsoft.XLANGs.RuntimeTypes.EventData( Microsoft.XLANGs.RuntimeTypes.Operation.End | Microsoft.XLANGs.RuntimeTypes.Operation.If),
            new Microsoft.XLANGs.RuntimeTypes.EventData( Microsoft.XLANGs.RuntimeTypes.Operation.Start | Microsoft.XLANGs.RuntimeTypes.Operation.Send),
            new Microsoft.XLANGs.RuntimeTypes.EventData( Microsoft.XLANGs.RuntimeTypes.Operation.Start | Microsoft.XLANGs.RuntimeTypes.Operation.Catch),
            new Microsoft.XLANGs.RuntimeTypes.EventData( Microsoft.XLANGs.RuntimeTypes.Operation.End | Microsoft.XLANGs.RuntimeTypes.Operation.Catch),
            new Microsoft.XLANGs.RuntimeTypes.EventData( Microsoft.XLANGs.RuntimeTypes.Operation.End | Microsoft.XLANGs.RuntimeTypes.Operation.Scope),
            new Microsoft.XLANGs.RuntimeTypes.EventData( Microsoft.XLANGs.RuntimeTypes.Operation.End | Microsoft.XLANGs.RuntimeTypes.Operation.Body)
        };

        public static int[] __progressLocation0 = new int[] { 0,0,0,3,3,};
        public static int[] __progressLocation1 = new int[] { 0,0,1,1,2,2,2,2,2,2,2,2,2,2,2,2,4,4,4,21,22,22,22,23,3,3,3,3,};
        public static int[] __progressLocation2 = new int[] { 6,6,6,7,7,7,7,7,7,7,7,7,9,9,10,10,10,10,10,10,10,10,10,10,10,10,10,10,10,10,10,10,10,10,10,10,10,10,10,10,11,11,11,12,13,13,14,15,15,15,15,15,15,15,15,15,16,16,16,16,};
        public static int[] __progressLocation3 = new int[] { 17,17,18,18,19,20,20,};

        public static int[][] __progressLocations = new int[4] [] {__progressLocation0,__progressLocation1,__progressLocation2,__progressLocation3};
        public override int[][] ProgressLocations {get {return __progressLocations;} }

        public Microsoft.XLANGs.Core.StopConditions segment0(Microsoft.XLANGs.Core.StopConditions stopOn)
        {
            Microsoft.XLANGs.Core.Segment __seg__ = _segments[0];
            Microsoft.XLANGs.Core.Context __ctx__ = (Microsoft.XLANGs.Core.Context)_stateMgrs[0];
            __MSCISOrchestration_root_0 __ctx0__ = (__MSCISOrchestration_root_0)_stateMgrs[0];
            __MSCISOrchestration_1 __ctx1__ = (__MSCISOrchestration_1)_stateMgrs[1];

            switch (__seg__.Progress)
            {
            case 0:
                MSCReceive_Port = new MSCReceive_PortType(0, this);
                SendRequest_Port = new SendRequest_PortType(1, this);
                __ctx__.PrologueCompleted = true;
                __ctx0__.__subWrapper0 = new Microsoft.XLANGs.Core.SubscriptionWrapper(ActivationSubGuids[0], MSCReceive_Port, this);
                if ( !PostProgressInc( __seg__, __ctx__, 1 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                if ((stopOn & Microsoft.XLANGs.Core.StopConditions.Initialized) != 0)
                    return Microsoft.XLANGs.Core.StopConditions.Initialized;
                goto case 1;
            case 1:
                __ctx1__ = new __MSCISOrchestration_1(this);
                _stateMgrs[1] = __ctx1__;
                if ( !PostProgressInc( __seg__, __ctx__, 2 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                goto case 2;
            case 2:
                __ctx0__.StartContext(__seg__, __ctx1__);
                if ( !PostProgressInc( __seg__, __ctx__, 3 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                return Microsoft.XLANGs.Core.StopConditions.Blocked;
            case 3:
                if (!__ctx0__.CleanupAndPrepareToCommit(__seg__))
                    return Microsoft.XLANGs.Core.StopConditions.Blocked;
                if ( !PostProgressInc( __seg__, __ctx__, 4 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                goto case 4;
            case 4:
                __ctx1__.Finally();
                ServiceDone(__seg__, (Microsoft.XLANGs.Core.Context)_stateMgrs[0]);
                __ctx0__.OnCommit();
                break;
            }
            return Microsoft.XLANGs.Core.StopConditions.Completed;
        }

        public Microsoft.XLANGs.Core.StopConditions segment1(Microsoft.XLANGs.Core.StopConditions stopOn)
        {
            Microsoft.XLANGs.Core.Envelope __msgEnv__ = null;
            Microsoft.XLANGs.Core.Segment __seg__ = _segments[1];
            Microsoft.XLANGs.Core.Context __ctx__ = (Microsoft.XLANGs.Core.Context)_stateMgrs[1];
            __MSCISOrchestration_root_0 __ctx0__ = (__MSCISOrchestration_root_0)_stateMgrs[0];
            __MSCISOrchestration_1 __ctx1__ = (__MSCISOrchestration_1)_stateMgrs[1];
            ____scope33_2 __ctx2__ = (____scope33_2)_stateMgrs[2];

            switch (__seg__.Progress)
            {
            case 0:
                __ctx1__.__requestId = null;
                __ctx1__.__serviceType = null;
                __ctx1__.__provider = null;
                __ctx1__.__requestSource = null;
                __ctx1__.__configMgr = null;
                __ctx1__.__messageConfig = null;
                __ctx1__.__requestMapName = null;
                __ctx1__.__responseMapName = null;
                __ctx1__.__requestMapType = null;
                __ctx1__.__responseMapType = null;
                __ctx1__.__ServiceType = null;
                __ctx1__.__Recipient = null;
                __ctx1__.__recipient = null;
                __ctx__.PrologueCompleted = true;
                if ( !PostProgressInc( __seg__, __ctx__, 1 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                goto case 1;
            case 1:
                if ( !PreProgressInc( __seg__, __ctx__, 2 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                Tracker.FireEvent(__eventLocations[0],__eventData[0],_stateMgrs[1].TrackDataStream );
                if (IsDebugged)
                    return Microsoft.XLANGs.Core.StopConditions.InBreakpoint;
                goto case 2;
            case 2:
                if ( !PreProgressInc( __seg__, __ctx__, 3 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                Tracker.FireEvent(__eventLocations[1],__eventData[1],_stateMgrs[1].TrackDataStream );
                if (IsDebugged)
                    return Microsoft.XLANGs.Core.StopConditions.InBreakpoint;
                goto case 3;
            case 3:
                if (!MSCReceive_Port.GetMessageId(__ctx0__.__subWrapper0.getSubscription(this), __seg__, __ctx1__, out __msgEnv__))
                    return Microsoft.XLANGs.Core.StopConditions.Blocked;
                if (__ctx1__.__MSCMessage != null)
                    __ctx1__.UnrefMessage(__ctx1__.__MSCMessage);
                __ctx1__.__MSCMessage = new __messagetype_MSC_IntegrationService_Schemas_cdiInterop("MSCMessage", __ctx1__);
                __ctx1__.RefMessage(__ctx1__.__MSCMessage);
                MSCReceive_Port.ReceiveMessage(0, __msgEnv__, __ctx1__.__MSCMessage, null, (Microsoft.XLANGs.Core.Context)_stateMgrs[1], __seg__);
                if ( !PostProgressInc( __seg__, __ctx__, 4 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                goto case 4;
            case 4:
                if ( !PreProgressInc( __seg__, __ctx__, 5 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                {
                    Microsoft.XLANGs.RuntimeTypes.EventData __edata = new Microsoft.XLANGs.RuntimeTypes.EventData(Microsoft.XLANGs.RuntimeTypes.Operation.End | Microsoft.XLANGs.RuntimeTypes.Operation.Receive);
                    __edata.Messages.Add(__ctx1__.__MSCMessage);
                    __edata.PortName = @"MSCReceive_Port";
                    Tracker.FireEvent(__eventLocations[2],__edata,_stateMgrs[1].TrackDataStream );
                }
                if (IsDebugged)
                    return Microsoft.XLANGs.Core.StopConditions.InBreakpoint;
                goto case 5;
            case 5:
                __ctx1__.__requestId = "";
                RootService.CommitStateManager.UserCodeCalled = true;
                if ( !PostProgressInc( __seg__, __ctx__, 6 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                goto case 6;
            case 6:
                __ctx1__.__serviceType = "";
                if ( !PostProgressInc( __seg__, __ctx__, 7 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                goto case 7;
            case 7:
                __ctx1__.__provider = "";
                if ( !PostProgressInc( __seg__, __ctx__, 8 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                goto case 8;
            case 8:
                __ctx1__.__requestSource = "";
                if ( !PostProgressInc( __seg__, __ctx__, 9 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                goto case 9;
            case 9:
                __ctx1__.__configMgr = new MSC.IntegrationService.ConfigService.ConfigManager();
                RootService.CommitStateManager.UserCodeCalled = true;
                if ( !PostProgressInc( __seg__, __ctx__, 10 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                goto case 10;
            case 10:
                __ctx1__.__messageConfig = new MSC.IntegrationService.ConfigService.MessageConfig();
                RootService.CommitStateManager.UserCodeCalled = true;
                if ( !PostProgressInc( __seg__, __ctx__, 11 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                goto case 11;
            case 11:
                __ctx1__.__requestMapName = "";
                if (__ctx1__ != null)
                    __ctx1__.__requestMapName = null;
                if ( !PostProgressInc( __seg__, __ctx__, 12 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                goto case 12;
            case 12:
                __ctx1__.__responseMapName = "";
                if (__ctx1__ != null)
                    __ctx1__.__responseMapName = null;
                if ( !PostProgressInc( __seg__, __ctx__, 13 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                goto case 13;
            case 13:
                __ctx1__.__ServiceType = new MSC.IntegrationService.ConfigService.ServiceType();
                RootService.CommitStateManager.UserCodeCalled = true;
                if ( !PostProgressInc( __seg__, __ctx__, 14 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                goto case 14;
            case 14:
                __ctx1__.__Recipient = new MSC.IntegrationService.ConfigService.Recipient();
                RootService.CommitStateManager.UserCodeCalled = true;
                if ( !PostProgressInc( __seg__, __ctx__, 15 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                goto case 15;
            case 15:
                __ctx1__.__recipient = "";
                if ( !PostProgressInc( __seg__, __ctx__, 16 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                goto case 16;
            case 16:
                if ( !PreProgressInc( __seg__, __ctx__, 17 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                Tracker.FireEvent(__eventLocations[4],__eventData[2],_stateMgrs[1].TrackDataStream );
                if (IsDebugged)
                    return Microsoft.XLANGs.Core.StopConditions.InBreakpoint;
                goto case 17;
            case 17:
                __ctx2__ = new ____scope33_2(this);
                _stateMgrs[2] = __ctx2__;
                if ( !PostProgressInc( __seg__, __ctx__, 18 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                goto case 18;
            case 18:
                __ctx1__.StartContext(__seg__, __ctx2__);
                if ( !PostProgressInc( __seg__, __ctx__, 19 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                return Microsoft.XLANGs.Core.StopConditions.Blocked;
            case 19:
                if ( !PreProgressInc( __seg__, __ctx__, 20 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                if (__ctx1__ != null)
                    __ctx1__.__recipient = null;
                if (__ctx1__ != null)
                    __ctx1__.__Recipient = null;
                if (__ctx1__ != null)
                    __ctx1__.__ServiceType = null;
                if (__ctx1__ != null)
                    __ctx1__.__messageConfig = null;
                if (__ctx1__ != null)
                    __ctx1__.__configMgr = null;
                if (__ctx1__ != null)
                    __ctx1__.__requestSource = null;
                if (__ctx1__ != null)
                    __ctx1__.__provider = null;
                if (__ctx1__ != null)
                    __ctx1__.__serviceType = null;
                if (__ctx1__ != null)
                    __ctx1__.__requestId = null;
                if (__ctx1__ != null && __ctx1__.__MSCMessage != null)
                {
                    __ctx1__.UnrefMessage(__ctx1__.__MSCMessage);
                    __ctx1__.__MSCMessage = null;
                }
                if (SendRequest_Port != null)
                {
                    SendRequest_Port.Close(__ctx1__, __seg__);
                    SendRequest_Port = null;
                }
                Tracker.FireEvent(__eventLocations[21],__eventData[11],_stateMgrs[1].TrackDataStream );
                __ctx2__.Finally();
                if (IsDebugged)
                    return Microsoft.XLANGs.Core.StopConditions.InBreakpoint;
                goto case 20;
            case 20:
                if ( !PreProgressInc( __seg__, __ctx__, 21 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                Tracker.FireEvent(__eventLocations[22],__eventData[8],_stateMgrs[1].TrackDataStream );
                if (IsDebugged)
                    return Microsoft.XLANGs.Core.StopConditions.InBreakpoint;
                goto case 21;
            case 21:
                if (!__ctx1__.PrepareToPendingCommit(__seg__))
                    return Microsoft.XLANGs.Core.StopConditions.Blocked;
                if ( !PostProgressInc( __seg__, __ctx__, 22 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                goto case 22;
            case 22:
                if ( !PreProgressInc( __seg__, __ctx__, 23 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                MSCReceive_Port.SendMessage(0, __ctx1__.__MSCResponseMsg, null, null, __ctx1__, __seg__ , Microsoft.XLANGs.Core.ActivityFlags.NextActivityPersists );
                if (MSCReceive_Port != null)
                {
                    MSCReceive_Port.Close(__ctx1__, __seg__);
                    MSCReceive_Port = null;
                }
                if ((stopOn & Microsoft.XLANGs.Core.StopConditions.OutgoingResp) != 0)
                    return Microsoft.XLANGs.Core.StopConditions.OutgoingResp;
                goto case 23;
            case 23:
                if ( !PreProgressInc( __seg__, __ctx__, 24 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                {
                    Microsoft.XLANGs.RuntimeTypes.EventData __edata = new Microsoft.XLANGs.RuntimeTypes.EventData(Microsoft.XLANGs.RuntimeTypes.Operation.End | Microsoft.XLANGs.RuntimeTypes.Operation.Send);
                    __edata.Messages.Add(__ctx1__.__MSCResponseMsg);
                    __edata.PortName = @"MSCReceive_Port";
                    Tracker.FireEvent(__eventLocations[23],__edata,_stateMgrs[1].TrackDataStream );
                }
                if (__ctx1__ != null && __ctx1__.__MSCResponseMsg != null)
                {
                    __ctx1__.UnrefMessage(__ctx1__.__MSCResponseMsg);
                    __ctx1__.__MSCResponseMsg = null;
                }
                if (IsDebugged)
                    return Microsoft.XLANGs.Core.StopConditions.InBreakpoint;
                goto case 24;
            case 24:
                if ( !PreProgressInc( __seg__, __ctx__, 25 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                Tracker.FireEvent(__eventLocations[3],__eventData[12],_stateMgrs[1].TrackDataStream );
                if (IsDebugged)
                    return Microsoft.XLANGs.Core.StopConditions.InBreakpoint;
                goto case 25;
            case 25:
                if (!__ctx1__.CleanupAndPrepareToCommit(__seg__))
                    return Microsoft.XLANGs.Core.StopConditions.Blocked;
                if ( !PostProgressInc( __seg__, __ctx__, 26 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                goto case 26;
            case 26:
                if ( !PreProgressInc( __seg__, __ctx__, 27 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                __ctx1__.OnCommit();
                goto case 27;
            case 27:
                __seg__.SegmentDone();
                _segments[0].PredecessorDone(this);
                break;
            }
            return Microsoft.XLANGs.Core.StopConditions.Completed;
        }

        public Microsoft.XLANGs.Core.StopConditions segment2(Microsoft.XLANGs.Core.StopConditions stopOn)
        {
            Microsoft.XLANGs.Core.Envelope __msgEnv__ = null;
            bool __condition__;
            Microsoft.XLANGs.Core.Segment __seg__ = _segments[2];
            Microsoft.XLANGs.Core.Context __ctx__ = (Microsoft.XLANGs.Core.Context)_stateMgrs[2];
            __MSCISOrchestration_root_0 __ctx0__ = (__MSCISOrchestration_root_0)_stateMgrs[0];
            __MSCISOrchestration_1 __ctx1__ = (__MSCISOrchestration_1)_stateMgrs[1];
            ____scope33_2 __ctx2__ = (____scope33_2)_stateMgrs[2];

            switch (__seg__.Progress)
            {
            case 0:
                __ctx__.PrologueCompleted = true;
                if ( !PostProgressInc( __seg__, __ctx__, 1 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                goto case 1;
            case 1:
                if ( !PreProgressInc( __seg__, __ctx__, 2 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                Tracker.FireEvent(__eventLocations[6],__eventData[3],_stateMgrs[2].TrackDataStream );
                if (IsDebugged)
                    return Microsoft.XLANGs.Core.StopConditions.InBreakpoint;
                goto case 2;
            case 2:
                System.Diagnostics.EventLog.WriteEntry("MSCIS", "Received request, assigning variables");
                RootService.CommitStateManager.UserCodeCalled = true;
                if ( !PostProgressInc( __seg__, __ctx__, 3 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                goto case 3;
            case 3:
                if ( !PreProgressInc( __seg__, __ctx__, 4 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                Tracker.FireEvent(__eventLocations[7],__eventData[4],_stateMgrs[2].TrackDataStream );
                if (IsDebugged)
                    return Microsoft.XLANGs.Core.StopConditions.InBreakpoint;
                goto case 4;
            case 4:
                __ctx1__.__requestId = (System.String)__ctx1__.__MSCMessage.GetPropertyValueThrows(typeof(MSC.IntegrationService.Schemas.PropertySchema.requestId));
                RootService.CommitStateManager.UserCodeCalled = true;
                if ( !PostProgressInc( __seg__, __ctx__, 5 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                goto case 5;
            case 5:
                __ctx1__.__serviceType = (System.String)__ctx1__.__MSCMessage.GetPropertyValueThrows(typeof(MSC.IntegrationService.Schemas.PropertySchema.serviceType));
                RootService.CommitStateManager.UserCodeCalled = true;
                if ( !PostProgressInc( __seg__, __ctx__, 6 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                goto case 6;
            case 6:
                __ctx1__.__requestSource = (System.String)__ctx1__.__MSCMessage.GetPropertyValueThrows(typeof(MSC.IntegrationService.Schemas.PropertySchema.requestSource));
                RootService.CommitStateManager.UserCodeCalled = true;
                if ( !PostProgressInc( __seg__, __ctx__, 7 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                goto case 7;
            case 7:
                __ctx1__.__provider = (System.String)__ctx1__.__MSCMessage.GetPropertyValueThrows(typeof(MSC.IntegrationService.Schemas.PropertySchema.provider));
                RootService.CommitStateManager.UserCodeCalled = true;
                if ( !PostProgressInc( __seg__, __ctx__, 8 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                goto case 8;
            case 8:
                __ctx1__.__recipient = (System.String)__ctx1__.__MSCMessage.GetPropertyValueThrows(typeof(MSC.IntegrationService.Schemas.PropertySchema.recipient));
                RootService.CommitStateManager.UserCodeCalled = true;
                if ( !PostProgressInc( __seg__, __ctx__, 9 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                goto case 9;
            case 9:
                __ctx1__.__messageConfig = __ctx1__.__configMgr.GetMessageConfig(__ctx1__.__provider);
                RootService.CommitStateManager.UserCodeCalled = true;
                if ( !PostProgressInc( __seg__, __ctx__, 10 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                goto case 10;
            case 10:
                __ctx1__.__ServiceType = __ctx1__.__messageConfig.ServiceTypes.GetServicType(__ctx1__.__serviceType);
                RootService.CommitStateManager.UserCodeCalled = true;
                if ( !PostProgressInc( __seg__, __ctx__, 11 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                goto case 11;
            case 11:
                __ctx1__.__Recipient = __ctx1__.__ServiceType.RecipientList.GetRecipient(__ctx1__.__recipient);
                RootService.CommitStateManager.UserCodeCalled = true;
                if ( !PostProgressInc( __seg__, __ctx__, 12 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                goto case 12;
            case 12:
                if ( !PreProgressInc( __seg__, __ctx__, 13 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                Tracker.FireEvent(__eventLocations[9],__eventData[3],_stateMgrs[2].TrackDataStream );
                if (IsDebugged)
                    return Microsoft.XLANGs.Core.StopConditions.InBreakpoint;
                goto case 13;
            case 13:
                System.Diagnostics.EventLog.WriteEntry("MSCIS", "In Construct_RequestMsg");
                RootService.CommitStateManager.UserCodeCalled = true;
                if ( !PostProgressInc( __seg__, __ctx__, 14 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                goto case 14;
            case 14:
                if ( !PreProgressInc( __seg__, __ctx__, 15 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                Tracker.FireEvent(__eventLocations[10],__eventData[4],_stateMgrs[2].TrackDataStream );
                if (IsDebugged)
                    return Microsoft.XLANGs.Core.StopConditions.InBreakpoint;
                goto case 15;
            case 15:
                __ctx1__.__requestMapType = System.Type.GetType(__ctx1__.__Recipient.RequestMapType);
                RootService.CommitStateManager.UserCodeCalled = true;
                if ( !PostProgressInc( __seg__, __ctx__, 16 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                goto case 16;
            case 16:
                if ( !PreProgressInc( __seg__, __ctx__, 17 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                Tracker.FireEvent(__eventLocations[5],__eventData[5],_stateMgrs[2].TrackDataStream );
                if (IsDebugged)
                    return Microsoft.XLANGs.Core.StopConditions.InBreakpoint;
                goto case 17;
            case 17:
                __condition__ = __ctx1__.__Recipient.Protocol == "SOAP";
                if (!__condition__)
                {
                    if ( !PostProgressInc( __seg__, __ctx__, 22 ) )
                        return Microsoft.XLANGs.Core.StopConditions.Paused;
                    goto case 22;
                }
                if ( !PostProgressInc( __seg__, __ctx__, 18 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                goto case 18;
            case 18:
                if ( !PreProgressInc( __seg__, __ctx__, 19 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                Tracker.FireEvent(__eventLocations[5],__eventData[6],_stateMgrs[2].TrackDataStream );
                if (IsDebugged)
                    return Microsoft.XLANGs.Core.StopConditions.InBreakpoint;
                goto case 19;
            case 19:
                {
                    __messagetype_System_Xml_XmlDocument __requestMessage = new __messagetype_System_Xml_XmlDocument("requestMessage", __ctx1__);

                    System.Diagnostics.EventLog.WriteEntry("MSCIS", "SOAP protocol, assigning parms.");
                    RootService.CommitStateManager.UserCodeCalled = true;
                    ApplyTransform(__ctx1__.__requestMapType, new object[] {__requestMessage.part}, new object[] {__ctx1__.__MSCMessage.part});
                    __requestMessage.SetPropertyValue(typeof(BTS.OutboundTransportType), "SOAP");
                    RootService.CommitStateManager.UserCodeCalled = true;
                    __requestMessage.SetPropertyValue(typeof(SOAP.AssemblyName), __ctx1__.__Recipient.Assembly);
                    RootService.CommitStateManager.UserCodeCalled = true;
                    __requestMessage.SetPropertyValue(typeof(SOAP.MethodName), __ctx1__.__Recipient.Operation);
                    RootService.CommitStateManager.UserCodeCalled = true;
                    __requestMessage.SetPropertyValue(typeof(SOAP.Username), __ctx1__.__Recipient.UserId);
                    RootService.CommitStateManager.UserCodeCalled = true;
                    __requestMessage.SetPropertyValue(typeof(SOAP.Password), __ctx1__.__Recipient.Password);
                    RootService.CommitStateManager.UserCodeCalled = true;

                    if (__ctx1__.__requestMessage != null)
                        __ctx1__.UnrefMessage(__ctx1__.__requestMessage);
                    __ctx1__.__requestMessage = __requestMessage;
                    __ctx1__.RefMessage(__ctx1__.__requestMessage);
                }
                __ctx1__.__requestMessage.ConstructionCompleteEvent(true);
                if ( !PostProgressInc( __seg__, __ctx__, 20 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                goto case 20;
            case 20:
                if ( !PreProgressInc( __seg__, __ctx__, 21 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                {
                    Microsoft.XLANGs.RuntimeTypes.EventData __edata = new Microsoft.XLANGs.RuntimeTypes.EventData(Microsoft.XLANGs.RuntimeTypes.Operation.End | Microsoft.XLANGs.RuntimeTypes.Operation.Construct);
                    __edata.Messages.Add(__ctx1__.__requestMessage);
                    __edata.Messages.Add(__ctx1__.__MSCMessage);
                    Tracker.FireEvent(__eventLocations[8],__edata,_stateMgrs[2].TrackDataStream );
                }
                if (IsDebugged)
                    return Microsoft.XLANGs.Core.StopConditions.InBreakpoint;
                goto case 21;
            case 21:
                if ( !PostProgressInc( __seg__, __ctx__, 39 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                goto case 39;
            case 22:
                if ( !PreProgressInc( __seg__, __ctx__, 23 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                Tracker.FireEvent(__eventLocations[5],__eventData[5],_stateMgrs[2].TrackDataStream );
                if (IsDebugged)
                    return Microsoft.XLANGs.Core.StopConditions.InBreakpoint;
                goto case 23;
            case 23:
                __condition__ = __ctx1__.__Recipient.Protocol == "WSE";
                if (!__condition__)
                {
                    if ( !PostProgressInc( __seg__, __ctx__, 28 ) )
                        return Microsoft.XLANGs.Core.StopConditions.Paused;
                    goto case 28;
                }
                if ( !PostProgressInc( __seg__, __ctx__, 24 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                goto case 24;
            case 24:
                if ( !PreProgressInc( __seg__, __ctx__, 25 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                Tracker.FireEvent(__eventLocations[5],__eventData[6],_stateMgrs[2].TrackDataStream );
                if (IsDebugged)
                    return Microsoft.XLANGs.Core.StopConditions.InBreakpoint;
                goto case 25;
            case 25:
                {
                    __messagetype_System_Xml_XmlDocument __requestMessage = new __messagetype_System_Xml_XmlDocument("requestMessage", __ctx1__);

                    System.Diagnostics.EventLog.WriteEntry("MSCIS", "WSE protocol, assigning parms.");
                    RootService.CommitStateManager.UserCodeCalled = true;
                    ApplyTransform(__ctx1__.__requestMapType, new object[] {__requestMessage.part}, new object[] {__ctx1__.__MSCMessage.part});
                    __requestMessage.SetPropertyValue(typeof(BTS.OutboundTransportType), "WSE");
                    RootService.CommitStateManager.UserCodeCalled = true;
                    __requestMessage.SetPropertyValue(typeof(WSE.SoapAction), "");
                    RootService.CommitStateManager.UserCodeCalled = true;
                    __requestMessage.SetPropertyValue(typeof(WSE.ConfidentialityUser), __ctx1__.__Recipient.UserId);
                    RootService.CommitStateManager.UserCodeCalled = true;
                    __requestMessage.SetPropertyValue(typeof(WSE.ConfidentialityPassword), __ctx1__.__Recipient.Password);
                    RootService.CommitStateManager.UserCodeCalled = true;
                    __requestMessage.SetPropertyValue(typeof(WSE.IntegrityPasswordOption), "Hashed");
                    RootService.CommitStateManager.UserCodeCalled = true;

                    if (__ctx1__.__requestMessage != null)
                        __ctx1__.UnrefMessage(__ctx1__.__requestMessage);
                    __ctx1__.__requestMessage = __requestMessage;
                    __ctx1__.RefMessage(__ctx1__.__requestMessage);
                }
                __ctx1__.__requestMessage.ConstructionCompleteEvent(true);
                if ( !PostProgressInc( __seg__, __ctx__, 26 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                goto case 26;
            case 26:
                if ( !PreProgressInc( __seg__, __ctx__, 27 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                {
                    Microsoft.XLANGs.RuntimeTypes.EventData __edata = new Microsoft.XLANGs.RuntimeTypes.EventData(Microsoft.XLANGs.RuntimeTypes.Operation.End | Microsoft.XLANGs.RuntimeTypes.Operation.Construct);
                    __edata.Messages.Add(__ctx1__.__requestMessage);
                    __edata.Messages.Add(__ctx1__.__MSCMessage);
                    Tracker.FireEvent(__eventLocations[8],__edata,_stateMgrs[2].TrackDataStream );
                }
                if (IsDebugged)
                    return Microsoft.XLANGs.Core.StopConditions.InBreakpoint;
                goto case 27;
            case 27:
                if ( !PostProgressInc( __seg__, __ctx__, 38 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                goto case 38;
            case 28:
                if ( !PreProgressInc( __seg__, __ctx__, 29 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                Tracker.FireEvent(__eventLocations[5],__eventData[5],_stateMgrs[2].TrackDataStream );
                if (IsDebugged)
                    return Microsoft.XLANGs.Core.StopConditions.InBreakpoint;
                goto case 29;
            case 29:
                __condition__ = __ctx1__.__Recipient.Protocol == "HTTP";
                if (!__condition__)
                {
                    if ( !PostProgressInc( __seg__, __ctx__, 34 ) )
                        return Microsoft.XLANGs.Core.StopConditions.Paused;
                    goto case 34;
                }
                if ( !PostProgressInc( __seg__, __ctx__, 30 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                goto case 30;
            case 30:
                if ( !PreProgressInc( __seg__, __ctx__, 31 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                Tracker.FireEvent(__eventLocations[5],__eventData[6],_stateMgrs[2].TrackDataStream );
                if (IsDebugged)
                    return Microsoft.XLANGs.Core.StopConditions.InBreakpoint;
                goto case 31;
            case 31:
                {
                    __messagetype_System_Xml_XmlDocument __requestMessage = new __messagetype_System_Xml_XmlDocument("requestMessage", __ctx1__);

                    System.Diagnostics.EventLog.WriteEntry("MSCIS", "HTTP protocol, assigning parms.");
                    RootService.CommitStateManager.UserCodeCalled = true;
                    ApplyTransform(__ctx1__.__requestMapType, new object[] {__requestMessage.part}, new object[] {__ctx1__.__MSCMessage.part});
                    __requestMessage.SetPropertyValue(typeof(BTS.OutboundTransportType), "HTTP");
                    RootService.CommitStateManager.UserCodeCalled = true;
                    __requestMessage.SetPropertyValue(typeof(MSC.IntegrationService.Schemas.PropertySchema.provider), __ctx1__.__provider);
                    RootService.CommitStateManager.UserCodeCalled = true;
                    __requestMessage.SetPropertyValue(typeof(MSC.IntegrationService.Schemas.PropertySchema.recipient), __ctx1__.__recipient);
                    RootService.CommitStateManager.UserCodeCalled = true;
                    __requestMessage.SetPropertyValue(typeof(MSC.IntegrationService.Schemas.PropertySchema.requestId), __ctx1__.__requestId);
                    RootService.CommitStateManager.UserCodeCalled = true;
                    __requestMessage.SetPropertyValue(typeof(MSC.IntegrationService.Schemas.PropertySchema.requestSource), __ctx1__.__requestSource);
                    RootService.CommitStateManager.UserCodeCalled = true;
                    __requestMessage.SetPropertyValue(typeof(MSC.IntegrationService.Schemas.PropertySchema.serviceType), __ctx1__.__serviceType);
                    RootService.CommitStateManager.UserCodeCalled = true;

                    if (__ctx1__.__requestMessage != null)
                        __ctx1__.UnrefMessage(__ctx1__.__requestMessage);
                    __ctx1__.__requestMessage = __requestMessage;
                    __ctx1__.RefMessage(__ctx1__.__requestMessage);
                }
                __ctx1__.__requestMessage.ConstructionCompleteEvent(true);
                if ( !PostProgressInc( __seg__, __ctx__, 32 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                goto case 32;
            case 32:
                if ( !PreProgressInc( __seg__, __ctx__, 33 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                {
                    Microsoft.XLANGs.RuntimeTypes.EventData __edata = new Microsoft.XLANGs.RuntimeTypes.EventData(Microsoft.XLANGs.RuntimeTypes.Operation.End | Microsoft.XLANGs.RuntimeTypes.Operation.Construct);
                    __edata.Messages.Add(__ctx1__.__requestMessage);
                    __edata.Messages.Add(__ctx1__.__MSCMessage);
                    Tracker.FireEvent(__eventLocations[8],__edata,_stateMgrs[2].TrackDataStream );
                }
                if (IsDebugged)
                    return Microsoft.XLANGs.Core.StopConditions.InBreakpoint;
                goto case 33;
            case 33:
                if ( !PostProgressInc( __seg__, __ctx__, 37 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                goto case 37;
            case 34:
                if ( !PreProgressInc( __seg__, __ctx__, 35 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                Tracker.FireEvent(__eventLocations[5],__eventData[6],_stateMgrs[2].TrackDataStream );
                if (IsDebugged)
                    return Microsoft.XLANGs.Core.StopConditions.InBreakpoint;
                goto case 35;
            case 35:
                {
                    __messagetype_System_Xml_XmlDocument __requestMessage = new __messagetype_System_Xml_XmlDocument("requestMessage", __ctx1__);

                    System.Diagnostics.EventLog.WriteEntry("MSCIS", "NO protocol.");
                    RootService.CommitStateManager.UserCodeCalled = true;
                    ApplyTransform(__ctx1__.__requestMapType, new object[] {__requestMessage.part}, new object[] {__ctx1__.__MSCMessage.part});

                    if (__ctx1__.__requestMessage != null)
                        __ctx1__.UnrefMessage(__ctx1__.__requestMessage);
                    __ctx1__.__requestMessage = __requestMessage;
                    __ctx1__.RefMessage(__ctx1__.__requestMessage);
                }
                __ctx1__.__requestMessage.ConstructionCompleteEvent(true);
                if ( !PostProgressInc( __seg__, __ctx__, 36 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                goto case 36;
            case 36:
                if ( !PreProgressInc( __seg__, __ctx__, 37 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                {
                    Microsoft.XLANGs.RuntimeTypes.EventData __edata = new Microsoft.XLANGs.RuntimeTypes.EventData(Microsoft.XLANGs.RuntimeTypes.Operation.End | Microsoft.XLANGs.RuntimeTypes.Operation.Construct);
                    __edata.Messages.Add(__ctx1__.__requestMessage);
                    __edata.Messages.Add(__ctx1__.__MSCMessage);
                    Tracker.FireEvent(__eventLocations[8],__edata,_stateMgrs[2].TrackDataStream );
                }
                if (IsDebugged)
                    return Microsoft.XLANGs.Core.StopConditions.InBreakpoint;
                goto case 37;
            case 37:
                if ( !PreProgressInc( __seg__, __ctx__, 38 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                Tracker.FireEvent(__eventLocations[8],__eventData[7],_stateMgrs[2].TrackDataStream );
                if (IsDebugged)
                    return Microsoft.XLANGs.Core.StopConditions.InBreakpoint;
                goto case 38;
            case 38:
                if ( !PreProgressInc( __seg__, __ctx__, 39 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                Tracker.FireEvent(__eventLocations[8],__eventData[7],_stateMgrs[2].TrackDataStream );
                if (IsDebugged)
                    return Microsoft.XLANGs.Core.StopConditions.InBreakpoint;
                goto case 39;
            case 39:
                if ( !PreProgressInc( __seg__, __ctx__, 40 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                if (__ctx1__ != null)
                    __ctx1__.__requestMapType = null;
                Tracker.FireEvent(__eventLocations[8],__eventData[7],_stateMgrs[2].TrackDataStream );
                if (IsDebugged)
                    return Microsoft.XLANGs.Core.StopConditions.InBreakpoint;
                goto case 40;
            case 40:
                if ( !PreProgressInc( __seg__, __ctx__, 41 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                Tracker.FireEvent(__eventLocations[11],__eventData[8],_stateMgrs[2].TrackDataStream );
                __ctx1__.__PromotedProp_Correlation = new Microsoft.XLANGs.Core.Correlation(this, 0, 0);
                if (IsDebugged)
                    return Microsoft.XLANGs.Core.StopConditions.InBreakpoint;
                goto case 41;
            case 41:
                if (!__ctx2__.PrepareToPendingCommit(__seg__))
                    return Microsoft.XLANGs.Core.StopConditions.Blocked;
                if ( !PostProgressInc( __seg__, __ctx__, 42 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                goto case 42;
            case 42:
                if ( !PreProgressInc( __seg__, __ctx__, 43 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                SendRequest_Port.SendMessage(0, __ctx1__.__requestMessage, new Microsoft.XLANGs.Core.Correlation[] { __ctx1__.__PromotedProp_Correlation }, null, out __ctx0__.__subWrapper1, __ctx2__, __seg__ );
                if (__ctx1__ != null && __ctx1__.__PromotedProp_Correlation != null)
                    __ctx1__.__PromotedProp_Correlation = null;
                if ((stopOn & Microsoft.XLANGs.Core.StopConditions.OutgoingRqst) != 0)
                    return Microsoft.XLANGs.Core.StopConditions.OutgoingRqst;
                goto case 43;
            case 43:
                if ( !PreProgressInc( __seg__, __ctx__, 44 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                {
                    Microsoft.XLANGs.RuntimeTypes.EventData __edata = new Microsoft.XLANGs.RuntimeTypes.EventData(Microsoft.XLANGs.RuntimeTypes.Operation.End | Microsoft.XLANGs.RuntimeTypes.Operation.Send);
                    __edata.Messages.Add(__ctx1__.__requestMessage);
                    __edata.PortName = @"SendRequest_Port";
                    Tracker.FireEvent(__eventLocations[12],__edata,_stateMgrs[2].TrackDataStream );
                }
                if (__ctx1__ != null && __ctx1__.__requestMessage != null)
                {
                    __ctx1__.UnrefMessage(__ctx1__.__requestMessage);
                    __ctx1__.__requestMessage = null;
                }
                if (IsDebugged)
                    return Microsoft.XLANGs.Core.StopConditions.InBreakpoint;
                goto case 44;
            case 44:
                if ( !PreProgressInc( __seg__, __ctx__, 45 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                Tracker.FireEvent(__eventLocations[13],__eventData[1],_stateMgrs[2].TrackDataStream );
                if (IsDebugged)
                    return Microsoft.XLANGs.Core.StopConditions.InBreakpoint;
                goto case 45;
            case 45:
                if (!SendRequest_Port.GetMessageId(__ctx0__.__subWrapper1.getSubscription(this), __seg__, __ctx1__, out __msgEnv__, _locations[0]))
                    return Microsoft.XLANGs.Core.StopConditions.Blocked;
                if (__ctx0__ != null && __ctx0__.__subWrapper1 != null)
                {
                    __ctx0__.__subWrapper1.Destroy(this, __ctx0__);
                    __ctx0__.__subWrapper1 = null;
                }
                if (__ctx1__.__responseMessage != null)
                    __ctx1__.UnrefMessage(__ctx1__.__responseMessage);
                __ctx1__.__responseMessage = new __messagetype_System_Xml_XmlDocument("responseMessage", __ctx1__);
                __ctx1__.RefMessage(__ctx1__.__responseMessage);
                SendRequest_Port.ReceiveMessage(0, __msgEnv__, __ctx1__.__responseMessage, null, (Microsoft.XLANGs.Core.Context)_stateMgrs[2], __seg__);
                if ( !PostProgressInc( __seg__, __ctx__, 46 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                goto case 46;
            case 46:
                if ( !PreProgressInc( __seg__, __ctx__, 47 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                {
                    Microsoft.XLANGs.RuntimeTypes.EventData __edata = new Microsoft.XLANGs.RuntimeTypes.EventData(Microsoft.XLANGs.RuntimeTypes.Operation.End | Microsoft.XLANGs.RuntimeTypes.Operation.Receive);
                    __edata.Messages.Add(__ctx1__.__responseMessage);
                    __edata.PortName = @"SendRequest_Port";
                    Tracker.FireEvent(__eventLocations[14],__edata,_stateMgrs[2].TrackDataStream );
                }
                if (IsDebugged)
                    return Microsoft.XLANGs.Core.StopConditions.InBreakpoint;
                goto case 47;
            case 47:
                if ( !PreProgressInc( __seg__, __ctx__, 48 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                Tracker.FireEvent(__eventLocations[15],__eventData[5],_stateMgrs[2].TrackDataStream );
                if (IsDebugged)
                    return Microsoft.XLANGs.Core.StopConditions.InBreakpoint;
                goto case 48;
            case 48:
                __condition__ = __ctx1__.__Recipient.ResponseMapType != "";
                if (!__condition__)
                {
                    if ( !PostProgressInc( __seg__, __ctx__, 53 ) )
                        return Microsoft.XLANGs.Core.StopConditions.Paused;
                    goto case 53;
                }
                if ( !PostProgressInc( __seg__, __ctx__, 49 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                goto case 49;
            case 49:
                if ( !PreProgressInc( __seg__, __ctx__, 50 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                Tracker.FireEvent(__eventLocations[5],__eventData[6],_stateMgrs[2].TrackDataStream );
                if (IsDebugged)
                    return Microsoft.XLANGs.Core.StopConditions.InBreakpoint;
                goto case 50;
            case 50:
                {
                    __messagetype_System_Xml_XmlDocument __MSCResponseMsg = new __messagetype_System_Xml_XmlDocument("MSCResponseMsg", __ctx1__);

                    __ctx1__.__responseMapType = System.Type.GetType(__ctx1__.__Recipient.ResponseMapType);
                    RootService.CommitStateManager.UserCodeCalled = true;
                    ApplyTransform(__ctx1__.__responseMapType, new object[] {__MSCResponseMsg.part}, new object[] {__ctx1__.__responseMessage.part});
                    if (__ctx1__ != null)
                        __ctx1__.__responseMapType = null;

                    if (__ctx1__.__MSCResponseMsg != null)
                        __ctx1__.UnrefMessage(__ctx1__.__MSCResponseMsg);
                    __ctx1__.__MSCResponseMsg = __MSCResponseMsg;
                    __ctx1__.RefMessage(__ctx1__.__MSCResponseMsg);
                }
                __ctx1__.__MSCResponseMsg.ConstructionCompleteEvent(true);
                if ( !PostProgressInc( __seg__, __ctx__, 51 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                goto case 51;
            case 51:
                if ( !PreProgressInc( __seg__, __ctx__, 52 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                {
                    Microsoft.XLANGs.RuntimeTypes.EventData __edata = new Microsoft.XLANGs.RuntimeTypes.EventData(Microsoft.XLANGs.RuntimeTypes.Operation.End | Microsoft.XLANGs.RuntimeTypes.Operation.Construct);
                    __edata.Messages.Add(__ctx1__.__MSCResponseMsg);
                    __edata.Messages.Add(__ctx1__.__responseMessage);
                    Tracker.FireEvent(__eventLocations[8],__edata,_stateMgrs[2].TrackDataStream );
                }
                if (IsDebugged)
                    return Microsoft.XLANGs.Core.StopConditions.InBreakpoint;
                goto case 52;
            case 52:
                if ( !PostProgressInc( __seg__, __ctx__, 56 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                goto case 56;
            case 53:
                if ( !PreProgressInc( __seg__, __ctx__, 54 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                Tracker.FireEvent(__eventLocations[5],__eventData[6],_stateMgrs[2].TrackDataStream );
                if (IsDebugged)
                    return Microsoft.XLANGs.Core.StopConditions.InBreakpoint;
                goto case 54;
            case 54:
                {
                    __messagetype_System_Xml_XmlDocument __MSCResponseMsg = new __messagetype_System_Xml_XmlDocument("MSCResponseMsg", __ctx1__);

                    __MSCResponseMsg.CopyFrom(__ctx1__.__responseMessage);
                    RootService.CommitStateManager.UserCodeCalled = true;

                    if (__ctx1__.__MSCResponseMsg != null)
                        __ctx1__.UnrefMessage(__ctx1__.__MSCResponseMsg);
                    __ctx1__.__MSCResponseMsg = __MSCResponseMsg;
                    __ctx1__.RefMessage(__ctx1__.__MSCResponseMsg);
                }
                __ctx1__.__MSCResponseMsg.ConstructionCompleteEvent(false);
                if ( !PostProgressInc( __seg__, __ctx__, 55 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                goto case 55;
            case 55:
                if ( !PreProgressInc( __seg__, __ctx__, 56 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                {
                    Microsoft.XLANGs.RuntimeTypes.EventData __edata = new Microsoft.XLANGs.RuntimeTypes.EventData(Microsoft.XLANGs.RuntimeTypes.Operation.End | Microsoft.XLANGs.RuntimeTypes.Operation.Construct);
                    __edata.Messages.Add(__ctx1__.__MSCResponseMsg);
                    Tracker.FireEvent(__eventLocations[8],__edata,_stateMgrs[2].TrackDataStream );
                }
                if (IsDebugged)
                    return Microsoft.XLANGs.Core.StopConditions.InBreakpoint;
                goto case 56;
            case 56:
                if ( !PreProgressInc( __seg__, __ctx__, 57 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                if (__ctx1__ != null && __ctx1__.__responseMessage != null)
                {
                    __ctx1__.UnrefMessage(__ctx1__.__responseMessage);
                    __ctx1__.__responseMessage = null;
                }
                Tracker.FireEvent(__eventLocations[16],__eventData[7],_stateMgrs[2].TrackDataStream );
                if (IsDebugged)
                    return Microsoft.XLANGs.Core.StopConditions.InBreakpoint;
                goto case 57;
            case 57:
                if (!__ctx2__.CleanupAndPrepareToCommit(__seg__))
                    return Microsoft.XLANGs.Core.StopConditions.Blocked;
                if ( !PostProgressInc( __seg__, __ctx__, 58 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                goto case 58;
            case 58:
                if ( !PreProgressInc( __seg__, __ctx__, 59 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                __ctx2__.OnCommit();
                goto case 59;
            case 59:
                __seg__.SegmentDone();
                _segments[1].PredecessorDone(this);
                break;
            }
            return Microsoft.XLANGs.Core.StopConditions.Completed;
        }

        public Microsoft.XLANGs.Core.StopConditions segment3(Microsoft.XLANGs.Core.StopConditions stopOn)
        {
            Microsoft.XLANGs.Core.Segment __seg__ = _segments[3];
            Microsoft.XLANGs.Core.Context __ctx__ = (Microsoft.XLANGs.Core.Context)_stateMgrs[2];
            __MSCISOrchestration_1 __ctx1__ = (__MSCISOrchestration_1)_stateMgrs[1];
            ____scope33_2 __ctx2__ = (____scope33_2)_stateMgrs[2];

            switch (__seg__.Progress)
            {
            case 0:
                OnBeginCatchHandler(2);
                if ( !PostProgressInc( __seg__, __ctx__, 1 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                goto case 1;
            case 1:
                if ( !PreProgressInc( __seg__, __ctx__, 2 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                Tracker.FireEvent(__eventLocations[17],__eventData[9],_stateMgrs[2].TrackDataStream );
                if (IsDebugged)
                    return Microsoft.XLANGs.Core.StopConditions.InBreakpoint;
                goto case 2;
            case 2:
                if ( !PreProgressInc( __seg__, __ctx__, 3 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                Tracker.FireEvent(__eventLocations[18],__eventData[6],_stateMgrs[2].TrackDataStream );
                if (IsDebugged)
                    return Microsoft.XLANGs.Core.StopConditions.InBreakpoint;
                goto case 3;
            case 3:
                {
                    __messagetype_System_Xml_XmlDocument __MSCResponseMsg = new __messagetype_System_Xml_XmlDocument("MSCResponseMsg", __ctx1__);

                    System.Diagnostics.EventLog.WriteEntry("MSCIS", "Error occurred:" + __ctx2__.__ex_0.Message);
                    RootService.CommitStateManager.UserCodeCalled = true;
                    __MSCResponseMsg.CopyFrom(__ctx1__.__MSCMessage);
                    RootService.CommitStateManager.UserCodeCalled = true;
                    __MSCResponseMsg.part.XPathStore(__ctx2__.__ex_0.Message, "/*[local-name()='clipboard']/*[local-name()='msgRequest']/*[local-name()='msgErrors']/*[local-name()='msgError']/@*[local-name()='message']");
                    RootService.CommitStateManager.UserCodeCalled = true;
                    if (__ctx2__ != null)
                        __ctx2__.__ex_0 = null;

                    if (__ctx1__.__MSCResponseMsg != null)
                        __ctx1__.UnrefMessage(__ctx1__.__MSCResponseMsg);
                    __ctx1__.__MSCResponseMsg = __MSCResponseMsg;
                    __ctx1__.RefMessage(__ctx1__.__MSCResponseMsg);
                }
                __ctx1__.__MSCResponseMsg.ConstructionCompleteEvent(false);
                if ( !PostProgressInc( __seg__, __ctx__, 4 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                goto case 4;
            case 4:
                if ( !PreProgressInc( __seg__, __ctx__, 5 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                {
                    Microsoft.XLANGs.RuntimeTypes.EventData __edata = new Microsoft.XLANGs.RuntimeTypes.EventData(Microsoft.XLANGs.RuntimeTypes.Operation.End | Microsoft.XLANGs.RuntimeTypes.Operation.Construct);
                    __edata.Messages.Add(__ctx1__.__MSCResponseMsg);
                    Tracker.FireEvent(__eventLocations[19],__edata,_stateMgrs[2].TrackDataStream );
                }
                if (IsDebugged)
                    return Microsoft.XLANGs.Core.StopConditions.InBreakpoint;
                goto case 5;
            case 5:
                if ( !PreProgressInc( __seg__, __ctx__, 6 ) )
                    return Microsoft.XLANGs.Core.StopConditions.Paused;
                Tracker.FireEvent(__eventLocations[20],__eventData[10],_stateMgrs[2].TrackDataStream );
                if (IsDebugged)
                    return Microsoft.XLANGs.Core.StopConditions.InBreakpoint;
                goto case 6;
            case 6:
                __ctx2__.__exv__ = null;
                OnEndCatchHandler(2, __seg__);
                __seg__.SegmentDone();
                break;
            }
            return Microsoft.XLANGs.Core.StopConditions.Completed;
        }
        private static Microsoft.XLANGs.Core.CachedObject[] _locations = new Microsoft.XLANGs.Core.CachedObject[] {
            new Microsoft.XLANGs.Core.CachedObject(new System.Guid("{2C541C50-96C0-4B7C-8A24-D672E011A774}"))
        };

    }

    [System.SerializableAttribute]
    sealed public class __MSC_IntegrationService_Schemas_cdiInterop__ : Microsoft.XLANGs.Core.XSDPart
    {
        private static MSC.IntegrationService.Schemas.cdiInterop _schema = new MSC.IntegrationService.Schemas.cdiInterop();

        public __MSC_IntegrationService_Schemas_cdiInterop__(Microsoft.XLANGs.Core.XMessage msg, string name, int index) : base(msg, name, index) { }

        
        #region part reflection support
        public static Microsoft.XLANGs.BaseTypes.SchemaBase PartSchema { get { return (Microsoft.XLANGs.BaseTypes.SchemaBase)_schema; } }
        #endregion // part reflection support
    }

    [Microsoft.XLANGs.BaseTypes.MessageTypeAttribute(
        Microsoft.XLANGs.BaseTypes.EXLangSAccess.ePublic,
        Microsoft.XLANGs.BaseTypes.EXLangSMessageInfo.eThirdKind,
        "MSC.IntegrationService.Schemas.cdiInterop",
        new System.Type[]{
            typeof(MSC.IntegrationService.Schemas.cdiInterop)
        },
        new string[]{
            "part"
        },
        new System.Type[]{
            typeof(__MSC_IntegrationService_Schemas_cdiInterop__)
        },
        0,
        @"http://mscanada.com#clipboard"
    )]
    [System.SerializableAttribute]
    sealed public class __messagetype_MSC_IntegrationService_Schemas_cdiInterop : Microsoft.BizTalk.XLANGs.BTXEngine.BTXMessage
    {
        public __MSC_IntegrationService_Schemas_cdiInterop__ part;

        private void __CreatePartWrappers()
        {
            part = new __MSC_IntegrationService_Schemas_cdiInterop__(this, "part", 0);
            this.AddPart("part", 0, part);
        }

        public __messagetype_MSC_IntegrationService_Schemas_cdiInterop(string msgName, Microsoft.XLANGs.Core.Context ctx) : base(msgName, ctx)
        {
            __CreatePartWrappers();
        }
    }

    [System.SerializableAttribute]
    sealed public class __Microsoft_XLANGs_BaseTypes_Any__ : Microsoft.XLANGs.Core.XSDPart
    {
        private static Microsoft.XLANGs.BaseTypes.Any _schema = new Microsoft.XLANGs.BaseTypes.Any();

        public __Microsoft_XLANGs_BaseTypes_Any__(Microsoft.XLANGs.Core.XMessage msg, string name, int index) : base(msg, name, index) { }

        
        #region part reflection support
        public static Microsoft.XLANGs.BaseTypes.SchemaBase PartSchema { get { return (Microsoft.XLANGs.BaseTypes.SchemaBase)_schema; } }
        #endregion // part reflection support
    }

    [Microsoft.XLANGs.BaseTypes.MessageTypeAttribute(
        Microsoft.XLANGs.BaseTypes.EXLangSAccess.ePublic,
        Microsoft.XLANGs.BaseTypes.EXLangSMessageInfo.eThirdKind,
        "System.Xml.XmlDocument",
        new System.Type[]{
            typeof(Microsoft.XLANGs.BaseTypes.Any)
        },
        new string[]{
            "part"
        },
        new System.Type[]{
            typeof(__Microsoft_XLANGs_BaseTypes_Any__)
        },
        0,
        Microsoft.XLANGs.Core.XMessage.AnyMessageTypeName
    )]
    [System.SerializableAttribute]
    sealed public class __messagetype_System_Xml_XmlDocument : Microsoft.BizTalk.XLANGs.BTXEngine.BTXMessage
    {
        public __Microsoft_XLANGs_BaseTypes_Any__ part;

        private void __CreatePartWrappers()
        {
            part = new __Microsoft_XLANGs_BaseTypes_Any__(this, "part", 0);
            this.AddPart("part", 0, part);
        }

        public __messagetype_System_Xml_XmlDocument(string msgName, Microsoft.XLANGs.Core.Context ctx) : base(msgName, ctx)
        {
            __CreatePartWrappers();
        }
    }

    [Microsoft.XLANGs.BaseTypes.BPELExportableAttribute(false)]
    sealed public class _MODULE_PROXY_ { }
}
