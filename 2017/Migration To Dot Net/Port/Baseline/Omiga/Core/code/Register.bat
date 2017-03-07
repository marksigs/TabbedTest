@echo off
ODI\ODIConverter\ReleaseUMinDependency\odiconverter -regserver
regsvr32 %1 "ExternalComponents\COM\omMP.dll"
regsvr32 %1 "ExternalComponents\COM\wPDF_X01.ocx"
regsvr32 %1 "Administration Interface\omAdmin\omAdmin.dll"
regsvr32 %1 "Audit\omAU\omAU.dll"
regsvr32 %1 "Base\omBase\omBase.dll"
regsvr32 %1 "CommonObjectGateway\omCOG\omCOG.dll"
regsvr32 %1 "CommonObjectGateway\omCOGMT\ReleaseUMinDependency\omCOGMT.dll"
regsvr32 %1 "CompensatingResourceManager\omCRM\omCRM.dll"
regsvr32 %1 "CompensatingResourceManager\omCRMLock\omCRMLock.dll"
regsvr32 %1 "dmsCompression\vc6\ReleaseUMinDependency\dmsCompression.dll"
regsvr32 %1 "HunterInterface\omHI\omHI.dll"
regsvr32 %1 "KFI PRINT SOLUTION\Delivery\Delivery.dll"
regsvr32 %1 "KFI PRINT SOLUTION\eKFI\eKFI.dll"
regsvr32 %1 "KFI PRINT SOLUTION\eKFIEngine\eKFIEngine.dll"
regsvr32 %1 "KFI PRINT SOLUTION\KFIPrintProcessor\KFIPrintProcessor.dll"
regsvr32 %1 "KFI PRINT SOLUTION\TemplateProcessor\TemplateProcessor.dll"
regsvr32 %1 MessageQueue\MessageQueue\omMQ\omMQ.dll
regsvr32 %1 MessageQueue\MessageQueue\omToMSMQ\omToMSMQ.dll
regsvr32 %1 MessageQueue\MessageQueue\omToNothing\omToNothing.dll
regsvr32 %1 MessageQueue\MessageQueue\omToOMMQ\omToOMMQ.dll
regsvr32 %1 MessageQueue\MessageQueueListener\MessageQueueComponentVC\ReleaseUMinDependencyW2K\MessageQueueComponentVC.dll
regsvr32 %1 omMutex\ReleaseUMinDependency\omMutex.dll
regsvr32 %1 omStream\ReleaseUMinDependency\omStream.dll
regsvr32 %1 Organisation\omOrg\omOrg.dll
regsvr32 %1 PAF\omPAF\omPAF.dll
regsvr32 %1 Printing\omDPS\omDPS.dll
regsvr32 %1 Printing\omFVS\omFVS.dll
regsvr32 %1 Printing\omPM\omPM.dll
regsvr32 %1 Printing\omPrint\omPrint.dll
regsvr32 %1 QuestInterface\omQuest\omQuest.dll
regsvr32 %1 RequestBroker\omRB\omRB.dll
regsvr32 %1 TaskManagement\msgTM\msgTM.dll
regsvr32 %1 TaskManagement\msgTM\TMBaseClient.dll
regsvr32 %1 TaskManagement\omTM\omTM.dll

