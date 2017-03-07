@echo off
SET CoreDir=C:\projects\MigrationToDotNet\Port\Baseline\Omiga\Core\code
SET ClientDir=C:\projects\MigrationToDotNet\Port\Baseline\Omiga\Client
SET InteropDir=C:\projects\MigrationToDotNet\Port\csharp\Interop

call mkInterop "%ClientDir%\Administration Interface\omAdmin\omAdmin.dll" %InteropDir%\Interop.omAdmin.dll omAdmin 
call mkInterop "%ClientDir%\Administration Interface\omAdminRules\omAdminRules.dll" %InteropDir%\Interop.omAdminRules.dll omAdminRules
call mkInterop "%ClientDir%\AgreementInPrinciple\omAIP\omAIP.dll" %InteropDir%\Interop.omAIP.dll omAIP
call mkInterop "%ClientDir%\Application\omApp\omApp.dll" %InteropDir%\Interop.omApp.dll omApp
call mkInterop "%ClientDir%\ApplicationProcessing\omAppProc\omAppProc.dll" %InteropDir%\Interop.omAppProc.dll omAppProc
call mkInterop "%ClientDir%\ApplicationProcessing\omAPRules\omAPRules.dll" %InteropDir%\Interop.omAPRules.dll omAPRules
call mkInterop "%ClientDir%\ApplicationQuote\omAQ\omAQ.dll" %InteropDir%\Interop.omAQ.dll omAQ
call mkInterop "%ClientDir%\BACS Interface\omBACS\omBACS.dll" %InteropDir%\Interop.omBACS.dll omBACS
call mkInterop "%ClientDir%\BatchProcessing\omBatch\omBatch.dll" %InteropDir%\Interop.omBatch.dll omBatch
call mkInterop "%ClientDir%\BuildingsAndContents\omBC\omBC.dll" %InteropDir%\Interop.omBC.dll omBC
call mkInterop "%ClientDir%\CaseTracking\omCaseTrack\omCaseTrack.dll" %InteropDir%\Interop.omCaseTrack.dll omCaseTrack
call mkInterop "%ClientDir%\CompletenessCheck\omComp\omComp.dll" %InteropDir%\Interop.omComp.dll omComp
call mkInterop "%ClientDir%\CompletenessCheck\omCompRules\omCompRules.dll" %InteropDir%\Interop.omCompRules.dll omCompRules
call mkInterop "%ClientDir%\CostModelling\omCM\omCM.dll" %InteropDir%\Interop.omCM.dll omCM
call mkInterop "%ClientDir%\CostModelling\omCMRules\omCMRules.dll" %InteropDir%\Interop.omCMRules.dll omCMRules
call mkInterop "%ClientDir%\CreateKFI\omCK\omCK.dll" %InteropDir%\Interop.omCK.dll omCK
call mkInterop "%ClientDir%\CreditCheck\omCC\omCC.dll" %InteropDir%\Interop.omCC.dll omCC
call mkInterop "%ClientDir%\CreditCheck\omExp\omExp.dll" %InteropDir%\Interop.omExp.dll omExp
call mkInterop "%ClientDir%\Customer\omCust\omCust.dll" %InteropDir%\Interop.omCust.dll omCust
call mkInterop "%ClientDir%\CustomerEmployment\omCE\omCE.dll" %InteropDir%\Interop.omCE.dll omCE
call mkInterop "%ClientDir%\CustomerFinancial\omCF\omCF.dll" %InteropDir%\Interop.omCF.dll omCF
call mkInterop "%ClientDir%\DataIngestion\omDataIngestion\omDataIngestion.dll" %InteropDir%\Interop.omDataIngestion.dll omDataIngestion
call mkInterop "%ClientDir%\DataIngestion\omDataIngestionRules\omDataIngestionRules.dll" %InteropDir%\Interop.omDataIngestionRules.dll omDataIngestionRules
call mkInterop "%ClientDir%\DecisionManager\omDM\omDM.dll" %InteropDir%\Interop.omDM.dll omDM
call mkInterop "%ClientDir%\HunterInterface\omHI\omHI.dll" %InteropDir%\Interop.omHI.dll omHI
call mkInterop "%ClientDir%\Import\omImp\omImp.dll" %InteropDir%\Interop.omImp.dll omImp
call mkInterop "%ClientDir%\Income Calcs\omIC\omIC.dll" %InteropDir%\Interop.omIC.dll omIC
call mkInterop "%ClientDir%\Ingestion\omIngestionManager\omIngestionManager.dll" %InteropDir%\Interop.omIngestionManager.dll omIngestionManager
call mkInterop "%ClientDir%\Ingestion\omIngestionRules\omIngestionRules.dll" %InteropDir%\Interop.omIngestionRules.dll omIngestionRules
call mkInterop "%ClientDir%\Intermediary\omIM\omIM.dll" %InteropDir%\Interop.omIM.dll omIM
call mkInterop "%ClientDir%\Introducer\omInt\omInt.dll" %InteropDir%\Interop.omInt.dll omInt
call mkInterop "%ClientDir%\KeyFactsIllustration\omKFIHelper\omKFIHelper.dll" %InteropDir%\Interop.omKFIHelper.dll omKFIHelper
call mkInterop "%ClientDir%\LifeCover\omLC\omLC.dll" %InteropDir%\Interop.omLC.dll omLC
call mkInterop "%ClientDir%\MortgageProduct\omMP\omMP.dll" %InteropDir%\Interop.omMP.dll omMP
call mkInterop "%ClientDir%\ODI\ODITransformer\ODITransformer.dll" %InteropDir%\Interop.ODITransformer.dll ODITransformer
call mkInterop "%ClientDir%\Omiga4ToOmiga3Download\om4to3\om4to3.dll" %InteropDir%\Interop.om4to3.dll om4to3
call mkInterop "%ClientDir%\Pack Manager\omPack\omPack.dll" %InteropDir%\Interop.omPack.dll omPack
call mkInterop "%ClientDir%\PaymentProcessing\omPayProc\omPayProc.dll" %InteropDir%\Interop.omPayProc.dll omPayProc
call mkInterop "%ClientDir%\PaymentProtection\omPP\omPP.dll" %InteropDir%\Interop.omPP.dll omPP
call mkInterop "%ClientDir%\Print Data Manager\omPDM\omPDM.dll" %InteropDir%\Interop.omPDM.dll omPDM
call mkInterop "%ClientDir%\QuickQuote\omQQ\omQQ.dll" %InteropDir%\Interop.omQQ.dll omQQ
call mkInterop "%ClientDir%\RateChange\omRC\omRC.dll" %InteropDir%\Interop.omRC.dll omRC
call mkInterop "%ClientDir%\RegisterDocumentToCase\omRegisterDocumentToCase\omRegisterDocumentToCase.dll" %InteropDir%\Interop.omRegisterDocumentToCase.dll omRegisterDocumentToCase
call mkInterop "%ClientDir%\ReportOnTitle\omROT\omROT.dll" %InteropDir%\Interop.omROT.dll omROT
call mkInterop "%ClientDir%\RequestBroker\omRB\omRB.dll" %InteropDir%\Interop.omRB.dll omRB
call mkInterop "%ClientDir%\RequestBroker\omRBlg\omRBlg.dll" %InteropDir%\Interop.omRBlg.dll omRBlg
call mkInterop "%ClientDir%\Risk Matrix\omRiskMatrix\omRiskMatrix.dll" %InteropDir%\Interop.omRiskMatrix.dll omRiskMatrix
call mkInterop "%ClientDir%\Risk Matrix\omRiskMatrixRules\RiskMatrixRules.dll" %InteropDir%\Interop.omRiskMatrixRules.dll omRiskMatrixRules
call mkInterop "%ClientDir%\RiskAssessment\omRA\omRA.dll" %InteropDir%\Interop.omRA.dll omRA
call mkInterop "%ClientDir%\RiskAssessment\omRARules\omRARules.dll" %InteropDir%\Interop.omRARules.dll omRARules
call mkInterop "%ClientDir%\Submission\omSub\omSub.dll" %InteropDir%\Interop.omSub.dll omSub
call mkInterop "%ClientDir%\TaskManagement\msgTM\msgTM.dll" %InteropDir%\Interop.msgTM.dll msgTM
call mkInterop "%ClientDir%\TaskManagement\msgTM\TmBaseClient.dll" %InteropDir%\Interop.TmBaseClient.dll TmBaseClient
call mkInterop "%ClientDir%\TaskManagement\omCDRules\omCDRules.dll" %InteropDir%\Interop.omCDRules.dll omCDRules
call mkInterop "%ClientDir%\TaskManagement\omTM\omTM.dll" %InteropDir%\Interop.omTM.dll omTM
call mkInterop "%ClientDir%\TaskManagement\omTmRules\omTmRules.dll" %InteropDir%\Interop.omTmRules.dll omTmRules
call mkInterop "%ClientDir%\ThirdParty\omTP\omTP.dll" %InteropDir%\Interop.omTP.dll omTP
call mkInterop "%ClientDir%\ValuationRules\omValuationRules\omValuationRules.dll" %InteropDir%\Interop.omValuationRules.dll omValuationRules
call mkInterop "%ClientDir%\WebServiceInterface\omWSInterface\omWSInterface.dll" %InteropDir%\Interop.omWSInterface.dll omWSInterface

:End
