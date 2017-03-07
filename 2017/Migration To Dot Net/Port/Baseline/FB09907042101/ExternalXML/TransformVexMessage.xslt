<?xml version="1.0" encoding="UTF-8"?>
<!-- 
==============================XML Document Control=============================

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Author		Date			Description
PE			10/11/2006	EP2_64 - Created / Xit2 Web Service
PE			27/11/2006	EP2_24
PE			07/12/2006	EP2_345 - Added InstructionSystemID
PE			31/01/2007	EP2_1029 - Various data mapping issues
PE			01/02/2007	EP2_997 - Set AppointmentDate to the correct value
PE			01/02/2007	EP_1029 - Set MainRoof to the correct value
PE			06/02/2007	EP_1029 - Set AppointmentDate and DateReceived nodes
PE			07/02/2007	EP2_1181 - Set Tenure node to "L" if LeaseHold is "YES"
PE			08/02/2007	EP2_1284 - Only return certification type for new builds.
PE			01/03/2007	EP2_1548 - Map MININGAREA and BUSINESSINBLOCKTYPEOFBUSINESS to correct values.
PE			15/03/2007	EP2_1741 - Truncate BUSINESSINBLOCKTYPEOFBUSINESS to 18 characters.
====================================================================================
  --> 
<xsl:stylesheet version="2.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:fn="http://www.w3.org/2005/xpath-functions">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<!--=================================================================================================-->
	<xsl:template match="/">
		<REQUEST>
			<VEX>
				<xsl:attribute name="VALSTATUS">
					<xsl:variable name="Status">
						<xsl:value-of select="/Instruction/Status/StatusCode"/>
					</xsl:variable>
					<xsl:choose>
						<xsl:when test="$Status='ACCEPT'"><xsl:value-of select="10"/></xsl:when>
						<xsl:when test="$Status='APPOINTMENT'"><xsl:value-of select="20"/></xsl:when>
						<xsl:when test="$Status='SUBMIT'"><xsl:value-of select="30"/></xsl:when>
						<xsl:when test="$Status='VALCANCEL'"><xsl:value-of select="40"/></xsl:when>
						<xsl:when test="$Status='VALONHOLD'"><xsl:value-of select="50"/></xsl:when>
						<xsl:when test="$Status='VALOFFHOLD'"><xsl:value-of select="60"/></xsl:when>
					</xsl:choose>
				</xsl:attribute>
				<xsl:attribute name="REASON"><xsl:value-of select="/Instruction/Status/Reason"/></xsl:attribute>
				<xsl:attribute name="APPOINTMENTDATE">
					<xsl:if test="/Instruction/Status/AppointmentDate">
						<xsl:value-of select="/Instruction/Status/AppointmentDate"/>	
					</xsl:if>					
				</xsl:attribute>
				<xsl:attribute name="SYSTEMID"><xsl:value-of select="/Instruction/Status/InstructionSystemID"/></xsl:attribute>
				<xsl:apply-templates select="/Instruction/Report"/>
				<xsl:apply-templates select="/Instruction/Report/SurveyDocuments/Document"/>
			</VEX>
			<APPLICATION>
				<xsl:attribute name="APPLICATIONNUMBER"><xsl:value-of select="/Instruction/InstructionRef"/></xsl:attribute>
			</APPLICATION>			
		</REQUEST>
	</xsl:template>
	<!--=================================================================================================-->
	<xsl:template match="Document">
		<PRINTDOCUMENT>
			<xsl:attribute name="DOCUMENTNAME">
				<xsl:value-of select="DocumentName"/>
			</xsl:attribute>
			<xsl:attribute name="DOCUMENTTYPE">
				<xsl:value-of select="DocumentType"/>
			</xsl:attribute>
			<xsl:attribute name="DOCUMENTCONTENTS">
				<xsl:value-of select="Document"/>
			</xsl:attribute>
		</PRINTDOCUMENT>			
	</xsl:template>
	<!--=================================================================================================-->
	<xsl:template match="Report">
		<!--=================================================================================================-->
			<VALUATION>
			<!--CREATEVALUATIONREPORTSUMMARY-->
			<xsl:attribute name="DATERECEIVED"></xsl:attribute>
			<xsl:attribute name="VALUERINVOICEAMOUNT"></xsl:attribute>
			<xsl:attribute name="VALFEEAUTHORISED"></xsl:attribute>
			<xsl:attribute name="VALUATIONSTATUS"></xsl:attribute>
		<!--=================================================================================================-->
			<!--CREATEVALUATIONREPORTVALUATION-->
			<xsl:attribute name="REINSTATEMENTVALUE">
				<xsl:value-of select="ReportData/MortgageValuationReport/Valuation/RebuildCost"/>
			</xsl:attribute>
			<xsl:attribute name="POSTWORKSVALUATION">
				<xsl:value-of select="ReportData/MortgageValuationReport/Valuation/ValueAfterEssentialRepairs"/>
			</xsl:attribute>
			<xsl:attribute name="RETENTIONSROADS"></xsl:attribute>
			<xsl:attribute name="PRESENTVALUATION">
				<xsl:value-of select="ReportData/MortgageValuationReport/Valuation/PresentValue"/>
			</xsl:attribute>
			<xsl:attribute name="RETENTIONWORKS">
				<xsl:value-of select="ReportData/MortgageValuationReport/Repairs/RecommendRetention"/>
			</xsl:attribute>
			<xsl:attribute name="CERTIFICATIONTYPE">
				<xsl:if test="ReportData/MortgageValuationReport/Construction/NewBuild='YES'">
				<xsl:value-of select="ReportData/MortgageValuationReport/NewBuild/CertificationType"/>
				</xsl:if>				
			</xsl:attribute>
			<xsl:attribute name="FUTUREINSPECTION"></xsl:attribute>
			<xsl:attribute name="GENERALREMARKS">
				<xsl:value-of select="substring(ReportData/MortgageValuationReport/ImportantInformation/ValuerRemarks,1,4000)"/>
			</xsl:attribute>
			<xsl:attribute name="REPAIRS"></xsl:attribute>
			<xsl:attribute name="GENERALREMARKNOTES"></xsl:attribute>
			<xsl:attribute name="REPAIRSNOTES">
				<xsl:variable name="notes">
					<xsl:value-of select="ReportData/MortgageValuationReport/Repairs/Repair1"/>
					<xsl:if test="ReportData/MortgageValuationReport/Repairs/Repair2!=''">
						<xsl:text>, </xsl:text>
						<xsl:value-of select="ReportData/MortgageValuationReport/Repairs/Repair2"/>
					</xsl:if>
					<xsl:if test="ReportData/MortgageValuationReport/Repairs/Repair3!=''">
						<xsl:text>, </xsl:text>
						<xsl:value-of select="ReportData/MortgageValuationReport/Repairs/Repair3"/>
					</xsl:if>
					<xsl:if test="ReportData/MortgageValuationReport/Repairs/Repair4!=''">
						<xsl:text>, </xsl:text>
						<xsl:value-of select="ReportData/MortgageValuationReport/Repairs/Repair4"/>
					</xsl:if>
					<xsl:if test="ReportData/MortgageValuationReport/Repairs/Repair5!=''">
						<xsl:text>, </xsl:text>
						<xsl:value-of select="ReportData/MortgageValuationReport/Repairs/Repair5"/>
					</xsl:if>
					<xsl:if test="ReportData/MortgageValuationReport/Repairs/Repair6!=''">
						<xsl:text>, </xsl:text>
						<xsl:value-of select="ReportData/MortgageValuationReport/Repairs/Repair6"/>
					</xsl:if>				
				</xsl:variable>
				<xsl:value-of select="substring($notes,1,4000)"/>
			</xsl:attribute>
			<xsl:attribute name="VALUATIONOK"></xsl:attribute>
			<xsl:attribute name="MULTIPLEOCCUPANCY"></xsl:attribute>
			<xsl:attribute name="REFERTOINSURANCE"></xsl:attribute>
			<xsl:attribute name="SIGNATURERETURNED">
				<xsl:choose>	
					<xsl:when test="ValuationValuer/Signatures/Signature/FullName!=''">1</xsl:when>
					<xsl:when test="ValuationValuer/Signatures/Signature/Title!=''">1</xsl:when>
					<xsl:when test="ValuationValuer/Signatures/Signature/ForeName!=''">1</xsl:when>
					<xsl:when test="ValuationValuer/Signatures/Signature/MiddleNames!=''">1</xsl:when>
					<xsl:when test="ValuationValuer/Signatures/Signature/LastName!=''">1</xsl:when>
					<xsl:when test="ValuationValuer/Signatures/Signature/Qualification!=''">1</xsl:when>
				</xsl:choose>
			</xsl:attribute>
			<xsl:attribute name="OVERALLCONDITION">
				<xsl:if test="ReportData/MortgageValuationReport/Saleability/SuitableSecurity='YES'">A</xsl:if>
				<xsl:if test="ReportData/MortgageValuationReport/Saleability/SuitableSecurity!='YES'">BA</xsl:if>						
			</xsl:attribute>
			<xsl:attribute name="SALEABILITY">
				<xsl:choose>
					<xsl:when test="ReportData/MortgageValuationReport/Saleability/PropertyDemand='GOOD'">AA</xsl:when>
					<xsl:when test="ReportData/MortgageValuationReport/Saleability/PropertyDemand='AVERAGE'">A</xsl:when>
					<xsl:when test="ReportData/MortgageValuationReport/Saleability/PropertyDemand='POOR'">BA</xsl:when>
				</xsl:choose>						
			</xsl:attribute>
		<!--=================================================================================================-->
			<!--CREATEVALUATIONREPORTPROPERTYRISKS-->
			<xsl:attribute name="LOCALSUBSIDENCE"></xsl:attribute>
			<xsl:attribute name="LIKELYTOAFFECTPROPERTY"></xsl:attribute>
			<xsl:attribute name="PROPERTYSUBSIDENCE"></xsl:attribute>
			<xsl:attribute name="LONGSTANDINGSUBIDENCE"></xsl:attribute>
			<xsl:attribute name="STRUCENGREPORTREQ"></xsl:attribute>
			<xsl:attribute name="AFFECTSALE"></xsl:attribute>
			<xsl:attribute name="TREEREPORTREQ"></xsl:attribute>
			<xsl:attribute name="NATUREOFRISKS"></xsl:attribute>
			<xsl:attribute name="TREETYPE"></xsl:attribute>
			<xsl:attribute name="TREEHEIGHT"></xsl:attribute>
			<xsl:attribute name="TREEDISTANCE"></xsl:attribute>
			<xsl:attribute name="WITHINCURTILEDGE"></xsl:attribute>
			<xsl:attribute name="MININGREPORTSREQ"></xsl:attribute>
			<xsl:attribute name="SPECIALISTREPORTREQ"></xsl:attribute>
			<xsl:attribute name="SPECIALISTREPORT"></xsl:attribute>
			<xsl:attribute name="PRONETOFLOODING"></xsl:attribute>
			<xsl:attribute name="TIMBERDAMPREPORTREQ"></xsl:attribute>
			<xsl:attribute name="ELECTICALREPORTREQ"></xsl:attribute>
			<xsl:attribute name="SOLICITORREFERENCEREQ"></xsl:attribute>
			<xsl:attribute name="SOLICITORNOTES"></xsl:attribute>
			<xsl:attribute name="OTHERHAZARDS"></xsl:attribute>
			<xsl:attribute name="ESSENTIALMATTERS"></xsl:attribute>
			<xsl:attribute name="ASBESTOSPOORCONDITION"></xsl:attribute>
			<xsl:attribute name="CAVITYWALLTIEFAILURE"></xsl:attribute>
			<xsl:attribute name="LARGEPANELSYSTEMAPPRAISED"></xsl:attribute>
			<xsl:attribute name="HISTORICBUILDINGREPAIRSREQ"></xsl:attribute>
			<xsl:attribute name="RETYPEIND"></xsl:attribute>
			<xsl:attribute name="RETYPEORIGINALLENDERNAME"></xsl:attribute>
			<xsl:attribute name="EXTENSIONSORALTERATIONS">
				<xsl:if test="ReportData/MortgageValuationReport/Alterations/PlanningPermiRequired='YES'">1</xsl:if>
				<xsl:if test="ReportData/MortgageValuationReport/Alterations/PlanningPermiRequired!='YES'">0</xsl:if>			
			</xsl:attribute>			
			<xsl:attribute name="NONRESIDENTIALLANDIND"></xsl:attribute>
			<xsl:attribute name="UNADOPTEDSHAREDACCESSISSUES"></xsl:attribute>
			<xsl:attribute name="TENANTEDPROPERTYIND">
				<xsl:if test="ReportData/MortgageValuationReport/Occupancy/PropertyTenanted='YES'">1</xsl:if>
				<xsl:if test="ReportData/MortgageValuationReport/Occupancy/PropertyTenanted!='YES'">0</xsl:if>			
			</xsl:attribute>			
			<xsl:attribute name="DEVELOPMENTPROPOSALS"></xsl:attribute>
			<xsl:attribute name="STRUCTURALISSUES">
				<xsl:if test="ReportData/MortgageValuationReport/Construction/StructuralIssues='YES'">1</xsl:if>
				<xsl:if test="ReportData/MortgageValuationReport/Construction/StructuralIssues!='YES'">0</xsl:if>			
			</xsl:attribute>
			<xsl:attribute name="NONPROGRESSIVESTRUCTURALISSUES">
				<xsl:if test="ReportData/MortgageValuationReport/Construction/LongstandingNonProgressive='YES'">1</xsl:if>
				<xsl:if test="ReportData/MortgageValuationReport/Construction/LongstandingNonProgressive!='YES'">0</xsl:if>			
			</xsl:attribute>			
			<xsl:attribute name="MUNDICRISK">
				<xsl:if test="ReportData/MortgageValuationReport/Construction/MundicProblems='YES'">1</xsl:if>
				<xsl:if test="ReportData/MortgageValuationReport/Construction/MundicProblems!='YES'">0</xsl:if>			
			</xsl:attribute>			
			<xsl:attribute name="RIGHTSOFWAY">
				<xsl:if test="ReportData/MortgageValuationReport/Legals/RightOfWay='YES'">1</xsl:if>
				<xsl:if test="ReportData/MortgageValuationReport/Legals/RightOfWay!='YES'">0</xsl:if>			
			</xsl:attribute>						
			<xsl:attribute name="SHAREDDRIVEORACCESS">
				<xsl:if test="ReportData/MortgageValuationReport/Legals/SharedAccess='YES'">1</xsl:if>
				<xsl:if test="ReportData/MortgageValuationReport/Legals/SharedAccess!='YES'">0</xsl:if>			
			</xsl:attribute>												
			<xsl:attribute name="GARAGEONSEPARATESITE">
				<xsl:if test="ReportData/MortgageValuationReport/Legals/SeparateGarage='YES'">1</xsl:if>
				<xsl:if test="ReportData/MortgageValuationReport/Legals/SeparateGarage!='YES'">0</xsl:if>			
			</xsl:attribute>				
			<xsl:attribute name="FLYINGFREEHOLD">
				<xsl:if test="ReportData/MortgageValuationReport/Legals/FlyFreeHold='YES'">1</xsl:if>
				<xsl:if test="ReportData/MortgageValuationReport/Legals/FlyFreeHold!='YES'">0</xsl:if>			
			</xsl:attribute>					
			<xsl:attribute name="FLYINGFREEHOLDPERCENT">
				<xsl:value-of select="ReportData/MortgageValuationReport/Legals/PercentageFlyFreeHold"/>
			</xsl:attribute>					
			<xsl:attribute name="SHAREDSERVICES">
				<xsl:if test="ReportData/MortgageValuationReport/Legals/SharedServices='YES'">1</xsl:if>
				<xsl:if test="ReportData/MortgageValuationReport/Legals/SharedServices!='YES'">0</xsl:if>			
			</xsl:attribute>				
			<xsl:attribute name="BOUNDARYILLDEFINED">
				<xsl:if test="ReportData/MortgageValuationReport/Legals/BoundaryIllDefined='YES'">1</xsl:if>
				<xsl:if test="ReportData/MortgageValuationReport/Legals/BoundaryIllDefined!='YES'">0</xsl:if>			
			</xsl:attribute>					
			<xsl:attribute name="LARGEPLOT">
				<xsl:if test="ReportData/MortgageValuationReport/Legals/PlotMoreThan10Acres='YES'">1</xsl:if>
				<xsl:if test="ReportData/MortgageValuationReport/Legals/PlotMoreThan10Acres!='YES'">0</xsl:if>			
			</xsl:attribute>			
			<xsl:attribute name="OTHERLEGALISSUES">
				<xsl:if test="ReportData/MortgageValuationReport/Legals/OtherLegals='YES'">1</xsl:if>
				<xsl:if test="ReportData/MortgageValuationReport/Legals/OtherLegals!='YES'">0</xsl:if>			
			</xsl:attribute>							
			<xsl:attribute name="SURVEYORTOINSPECTTITLE">
				<xsl:if test="ReportData/MortgageValuationReport/Legals/InspectTitlePlans='YES'">1</xsl:if>
				<xsl:if test="ReportData/MortgageValuationReport/Legals/InspectTitlePlans!='YES'">0</xsl:if>			
			</xsl:attribute>							
			<xsl:attribute name="MININGAREA">
				<xsl:if test="ReportData/MortgageValuationReport/Locality/MiningArea='YES'">1</xsl:if>
				<xsl:if test="ReportData/MortgageValuationReport/Locality/MiningArea!='YES'">0</xsl:if>			
			</xsl:attribute>							
			<xsl:attribute name="RADONGASAREA">
				<xsl:if test="ReportData/MortgageValuationReport/Locality/RadonArea='YES'">1</xsl:if>
				<xsl:if test="ReportData/MortgageValuationReport/Locality/RadonArea!='YES'">0</xsl:if>			
			</xsl:attribute>						
			<xsl:attribute name="RADONGASALLOWEDFOR">
				<xsl:if test="ReportData/MortgageValuationReport/Locality/RisksCalculatedIntoValuation='YES'">1</xsl:if>
				<xsl:if test="ReportData/MortgageValuationReport/Locality/RisksCalculatedIntoValuation!='YES'">0</xsl:if>			
			</xsl:attribute>						
			<xsl:attribute name="BOARDEDUPORDAMAGEDPROPERTY">
				<xsl:if test="ReportData/MortgageValuationReport/Locality/DamagedPropertiesInVicinity='YES'">1</xsl:if>
				<xsl:if test="ReportData/MortgageValuationReport/Locality/DamagedPropertiesInVicinity!='YES'">0</xsl:if>			
			</xsl:attribute>						
			<xsl:attribute name="OCCUPANCYRESTRICTIONS">
				<xsl:if test="ReportData/MortgageValuationReport/Occupancy/OccupancyRestrictions='YES'">1</xsl:if>
				<xsl:if test="ReportData/MortgageValuationReport/Occupancy/OccupancyRestrictions!='YES'">0</xsl:if>			
			</xsl:attribute>						
		<!--=================================================================================================-->
			<!--CREATEVALUATIONREPORTPROPERTYSERVICES-->
			<xsl:attribute name="NUMBEROFBEDROOMS"><xsl:value-of select="ReportData/MortgageValuationReport/Property/Bedroom"/></xsl:attribute>
			<xsl:attribute name="MAINSERVICES"></xsl:attribute>
			<xsl:attribute name="GAS"></xsl:attribute>
			<xsl:attribute name="WATER">
				<xsl:choose>
					<xsl:when test="ReportData/MortgageValuationReport/Services/MainsWater='YES'">M</xsl:when>
					<xsl:when test="ReportData/MortgageValuationReport/Services/PrivateWater='YES'">P</xsl:when>
				</xsl:choose>	
			</xsl:attribute>
			<xsl:attribute name="ELECTRICITY"></xsl:attribute>
			<xsl:attribute name="DRAINAGE">
				<xsl:choose>
					<xsl:when test="ReportData/MortgageValuationReport/Services/MainsDrainage='YES'">M</xsl:when>
					<xsl:when test="ReportData/MortgageValuationReport/Services/SepticTank='YES'">SN</xsl:when>
				</xsl:choose>	
			</xsl:attribute>
			<xsl:attribute name="HOTWATERCENTRALHEATING"></xsl:attribute>
			<xsl:attribute name="DEMANDINAREA">
				<xsl:choose>
					<xsl:when test="ReportData/MortgageValuationReport/Saleability/PropertyDemand='GOOD'">AA</xsl:when>
					<xsl:when test="ReportData/MortgageValuationReport/Saleability/PropertyDemand='AVERAGE'">A</xsl:when>
					<xsl:when test="ReportData/MortgageValuationReport/Saleability/PropertyDemand='POOR'">BA</xsl:when>
				</xsl:choose>			
			</xsl:attribute>
			<xsl:attribute name="OTHERFACTORS"></xsl:attribute>
			<xsl:attribute name="FUTURESALEABILITY">
				<xsl:if test="ReportData/MortgageValuationReport/Saleability/FutureSaleability='YES'">1</xsl:if>
				<xsl:if test="ReportData/MortgageValuationReport/Saleability/FutureSaleability!='YES'">0</xsl:if>									
			</xsl:attribute>
			<xsl:attribute name="ROADSMADEUPADOPTED">
				<xsl:if test="ReportData/MortgageValuationReport/Legals/RoadAdopted='YES'">1</xsl:if>
				<xsl:if test="ReportData/MortgageValuationReport/Legals/RoadAdopted!='YES'">0</xsl:if>						
			</xsl:attribute>
			<xsl:attribute name="ROADSTOBEADOPTED"></xsl:attribute>
			<xsl:attribute name="OTHERFACTORSNOTES"></xsl:attribute>
			<xsl:attribute name="HABITABLEROOMS"></xsl:attribute>
			<xsl:attribute name="LIVINGROOMS"><xsl:value-of select="ReportData/MortgageValuationReport/Property/LivingRoom"/></xsl:attribute>
			<xsl:attribute name="BATHROOMS"><xsl:value-of select="ReportData/MortgageValuationReport/Property/Bathroom"/></xsl:attribute>
			<xsl:attribute name="SEPERATEWCS"></xsl:attribute>
			<xsl:attribute name="GARAGES"><xsl:value-of select="ReportData/MortgageValuationReport/Property/Garage"/></xsl:attribute>
			<xsl:attribute name="PARKINGSPACES"><xsl:value-of select="ReportData/MortgageValuationReport/Property/ParkingSpace"/></xsl:attribute>
			<xsl:attribute name="CONSERVATORY"></xsl:attribute>
			<xsl:attribute name="GARDENINCLUDED"></xsl:attribute>
		<!--=================================================================================================-->
		<!--CREATEVALUATIONREPORTPROPERTYDETAILS-->
			<xsl:attribute name="TYPEOFPROPERTY">
					<xsl:choose>
						<xsl:when test="ReportData/MortgageValuationReport/Property/ConvertedStudioFlat='YES'">F</xsl:when>
						<xsl:when test="ReportData/MortgageValuationReport/Property/PurposeBuiltStudioFlat='YES'">F</xsl:when>
						<xsl:when test="ReportData/MortgageValuationReport/Property/ConvertedFlat='YES'">F</xsl:when>
						<xsl:when test="ReportData/MortgageValuationReport/Property/PurposeBuiltFlat='YES'">F</xsl:when>						
						<xsl:when test="ReportData/MortgageValuationReport/Property/Bungalow='YES'">B</xsl:when>												
						<xsl:when test="ReportData/MortgageValuationReport/Property/Maisonette='YES'">M</xsl:when>												
						<xsl:when test="ReportData/MortgageValuationReport/Property/Other='YES'">O</xsl:when>
						<xsl:when test="ReportData/MortgageValuationReport/Property/House='YES'">H</xsl:when>
					</xsl:choose>						
			</xsl:attribute>
			<xsl:attribute name="PROPERTYDESCRIPTION">
					<xsl:choose>
						<xsl:when test="ReportData/MortgageValuationReport/Property/PurposeBuiltFlat='YES'">F</xsl:when>
						<xsl:when test="ReportData/MortgageValuationReport/Property/ConvertedFlat='YES'">F</xsl:when>
						<xsl:when test="ReportData/MortgageValuationReport/Property/Maisonette='YES'">M</xsl:when>
						<xsl:when test="ReportData/MortgageValuationReport/Property/ConvertedStudioFlat='YES'">CS</xsl:when>
						<xsl:when test="ReportData/MortgageValuationReport/Property/PurposeBuiltStudioFlat='YES'">PBS</xsl:when>
						<xsl:when test="ReportData/MortgageValuationReport/Property/Detached='YES'">HD</xsl:when>
						<xsl:when test="ReportData/MortgageValuationReport/Property/SemiDetached='YES'">HS</xsl:when>
						<xsl:when test="ReportData/MortgageValuationReport/Property/MidTerrace='YES'">VM</xsl:when>
						<xsl:when test="ReportData/MortgageValuationReport/Property/EndTerrace='YES'">VE</xsl:when>
						<xsl:when test="ReportData/MortgageValuationReport/Property/BackToBack='YES'">BB</xsl:when>
					</xsl:choose>						
			</xsl:attribute>
			<xsl:attribute name="YEARBUILT"><xsl:value-of select="ReportData/MortgageValuationReport/Construction/YearOfConstruction"/></xsl:attribute>
			<xsl:attribute name="TENURE">
				<xsl:choose>
					<xsl:when test="ReportData/MortgageValuationReport/Tenure/FreeHold='YES'">F</xsl:when>
					<xsl:when test="ReportData/MortgageValuationReport/Tenure/LeaseHold='YES'">L</xsl:when>
					<xsl:when test="ReportData/MortgageValuationReport/Tenure/CommonHold='YES'">CH</xsl:when>
					<xsl:when test="ReportData/MortgageValuationReport/Tenure/ScottishFreeHold='YES'">FE</xsl:when>
				</xsl:choose>
			</xsl:attribute>
			<xsl:attribute name="POSTCODE"><xsl:value-of select="substring(NewAddress/PostCode,1,8)"/></xsl:attribute>
			<xsl:attribute name="FLATNUMBER"></xsl:attribute>
			<xsl:attribute name="BUILDINGORHOUSENAME"><xsl:value-of select="substring(NewAddress/HouseName,1,40)"/></xsl:attribute>
			<xsl:attribute name="BUILDINGORHOUSENUMBER"><xsl:value-of select="substring(NewAddress/HouseNumber,1,10)"/></xsl:attribute>
			<xsl:attribute name="STREET"><xsl:value-of select="substring(NewAddress/Street,1,40)"/></xsl:attribute>
			<xsl:attribute name="DISTRICT"><xsl:value-of select="substring(NewAddress/District,1,40)"/></xsl:attribute>
			<xsl:attribute name="TOWN"><xsl:value-of select="substring(NewAddress/City,1,40)"/></xsl:attribute>
			<xsl:attribute name="COUNTY"><xsl:value-of select="substring(NewAddress/County,1,40)"/></xsl:attribute>
			<xsl:attribute name="COUNTRY"><xsl:value-of select="NewAddress/Country"/></xsl:attribute>
			<xsl:attribute name="RESIDENCEAREA"><xsl:value-of select="ReportData/MortgageValuationReport/Construction/GrossExternalFloorArea"/></xsl:attribute>
			<xsl:attribute name="GARAGEAREA"></xsl:attribute>
			<xsl:attribute name="RESIDENTIALOUTBUILDINGSAREA"></xsl:attribute>
			<xsl:attribute name="AGRICULTURALOUTBUILDINGSAREA"></xsl:attribute>
			<xsl:attribute name="STRUCTURE">
				<xsl:if test="ReportData/MortgageValuationReport/Construction/WallsConventional='YES'">TC</xsl:if>
				<xsl:if test="ReportData/MortgageValuationReport/Construction/WallsConventional!='YES'">NTC</xsl:if>
			</xsl:attribute>
			<xsl:attribute name="MAINROOF">
				<xsl:if test="ReportData/MortgageValuationReport/Construction/RoofConventional='YES'">S</xsl:if>
				<xsl:if test="ReportData/MortgageValuationReport/Construction/RoofConventional='NO'">NTC</xsl:if>
			</xsl:attribute>
			<xsl:attribute name="GROUNDRENT"><xsl:value-of select="ReportData/MortgageValuationReport/Tenure/GroundRent"/></xsl:attribute>
			<xsl:attribute name="SERVICECHARGE"><xsl:value-of select="ReportData/MortgageValuationReport/Tenure/ServiceCharges"/></xsl:attribute>
			<xsl:attribute name="NUMBEROFKITCHENS"><xsl:value-of select="ReportData/MortgageValuationReport/Property/Kitchen"/></xsl:attribute>
			<xsl:attribute name="NUMBEROFOCCUPANTS"></xsl:attribute>
			<xsl:attribute name="CHIEFRENT"><xsl:value-of select="ReportData/MortgageValuationReport/Tenure/FeuDuty"/></xsl:attribute>
			<xsl:attribute name="RENTALINCOME"><xsl:value-of select="ReportData/MortgageValuationReport/BTL/RentUnFurnished"/></xsl:attribute>
			<xsl:attribute name="ESTROADCHARGE"></xsl:attribute>
			<xsl:attribute name="STOREYSINBLOCK"><xsl:value-of select="ReportData/MortgageValuationReport/Property/FloorsInBlock"/></xsl:attribute>
			<xsl:attribute name="FLOORINBLOCK"><xsl:value-of select="ReportData/MortgageValuationReport/Property/WhichFloor"/></xsl:attribute>
			<xsl:attribute name="BUSINESSINBLOCK">
				<xsl:if test="ReportData/MortgageValuationReport/Locality/AboveCommercialPrems='YES'">1</xsl:if>
				<xsl:if test="ReportData/MortgageValuationReport/Locality/AboveCommercialPrems!='YES'">0</xsl:if>						
			</xsl:attribute>
			<xsl:attribute name="INSTRUCTIONADDRESSCORRECT"></xsl:attribute>
			<xsl:attribute name="CONSTRUCTIONTYPEDETAILS"></xsl:attribute>
			<xsl:attribute name="ISSINGLESKINCONSTRUCTION"></xsl:attribute>
			<xsl:attribute name="ISSINGLESKINTWOSTOREY"></xsl:attribute>
			<xsl:attribute name="GENERALOBSERVATIONS"></xsl:attribute>
			<xsl:attribute name="EXLOCALAUTHORITY">
				<xsl:if test="ReportData/MortgageValuationReport/Locality/ExLocalAuthority='YES'">1</xsl:if>
				<xsl:if test="ReportData/MortgageValuationReport/Locality/ExLocalAuthority!='YES'">0</xsl:if>			
			</xsl:attribute>
			<xsl:attribute name="NONSTANDARDCONSTRUCTIONTYPE"></xsl:attribute>
			<xsl:attribute name="NEWPROPERTYINDICATOR">
				<xsl:if test="ReportData/MortgageValuationReport/Construction/NewBuild='YES'">1</xsl:if>
				<xsl:if test="ReportData/MortgageValuationReport/Construction/NewBuild!='YES'">0</xsl:if>
			</xsl:attribute>
			<xsl:attribute name="NUMBEROFUNITS"><xsl:value-of select="ReportData/MortgageValuationReport/Property/UnitsInDevelopment"/></xsl:attribute>
			<xsl:attribute name="FLOORS"><xsl:value-of select="ReportData/MortgageValuationReport/Property/Floor"/></xsl:attribute>
			<xsl:attribute name="BALCONYACCESS">
				<xsl:if test="ReportData/MortgageValuationReport/Property/PropertyViaBalcony='YES'">1</xsl:if>
				<xsl:if test="ReportData/MortgageValuationReport/Property/PropertyViaBalcony!='YES'">0</xsl:if>
			</xsl:attribute>			
			<xsl:attribute name="LIVEWORKUNITRESELEMENT">
				<xsl:if test="ReportData/MortgageValuationReport/Construction/LiveWorkUnit='YES'">1</xsl:if>
				<xsl:if test="ReportData/MortgageValuationReport/Construction/LiveWorkUnit!='YES'">0</xsl:if>
			</xsl:attribute>			
			<xsl:attribute name="COMMERCIALUSAGE">
				<xsl:if test="ReportData/MortgageValuationReport/Construction/CommercialUsage ='YES'">1</xsl:if>
				<xsl:if test="ReportData/MortgageValuationReport/Construction/CommercialUsage !='YES'">0</xsl:if>
			</xsl:attribute>												
			<xsl:attribute name="BUSINESSINBLOCKTYPEOFBUSINESS">
				<xsl:value-of select="substring(ReportData/MortgageValuationReport/Locality/CommercialUsage,1,18)"/>
			</xsl:attribute>												
			<xsl:attribute name="UNEXPIREDLEASE"><xsl:value-of select="ReportData/MortgageValuationReport/Tenure/LHUnexpiredTerm"/></xsl:attribute>
			<xsl:attribute name="EXLAPERCENTAGEPRIVATEOWNER"><xsl:value-of select="ReportData/MortgageValuationReport/Locality/PrivateOwnership"/></xsl:attribute>
			<xsl:attribute name="PPINCLSALESINCENTIVES">
				<xsl:if test="ReportData/MortgageValuationReport/NewBuild/IncludesSalesIncentives='YES'">1</xsl:if>
				<xsl:if test="ReportData/MortgageValuationReport/NewBuild/IncludesSalesIncentives!='YES'">0</xsl:if>
			</xsl:attribute>							
			<xsl:attribute name="RENTALDEMAND"><xsl:value-of select="ReportData/MortgageValuationReport/BTL/RentalDemand"/></xsl:attribute>						
			<xsl:attribute name="RENTABLECONDITION">
				<xsl:if test="ReportData/MortgageValuationReport/BTL/Rentable='YES'">1</xsl:if>
				<xsl:if test="ReportData/MortgageValuationReport/BTL/Rentable!='YES'">0</xsl:if>
			</xsl:attribute>							
		<!--=================================================================================================-->
		</VALUATION>
		<CONTACTDETAILS>
			<xsl:attribute name="CONTACTFORENAME">
				<xsl:value-of select="ValuationValuer/Signatures/Signature/ForeName"/>
				<xsl:if test="ValuationValuer/Signatures/Signature/MiddleNames!=''">
					<xsl:text> </xsl:text>
					<xsl:value-of select="ValuationValuer/Signatures/Signature/MiddleNames"/>
				</xsl:if>
			</xsl:attribute>
			<xsl:attribute name="CONTACTSURNAME">
				<xsl:value-of select="ValuationValuer/Signatures/Signature/LastName"/>
			</xsl:attribute>
			<xsl:attribute name="CONTACTTITLE">
				<xsl:value-of select="ValuationValuer/Signatures/Signature/Title"/>
			</xsl:attribute>
		</CONTACTDETAILS>
	</xsl:template>
	<!--=================================================================================================-->
</xsl:stylesheet>
