<?xml version="1.0" encoding="UTF-8"?>
<!--<?altova_samplexml C:\Dev\EPSOM\1 DEV Code\TaskManagement\omCDRules\cdTestAlias.xml?>-->

<!-- IK 27/04/2006 EP410 -->
<!-- LH 17/05/2006 EP558 -->
<!-- LH 13/06/2006 EP580 -->
<!-- IK 23/06/2006 EP767 -->
<!-- LH 16/02/2007 EP2_1281 -->
<!-- SR 19/03/2007 EP2_1828-->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:user="urn:my-scripts">
	<xsl:output method="xml" version="2.0" encoding="UTF-8" indent="no"/>
	<msxsl:script language="JavaScript" implements-prefix="user"><![CDATA[
		var hasTask = false;
		var datachanged_result = false;
		var datachanged = false;
				
		function SetDataChanged()
		{
			datachanged = true;
			datachanged_result = '1';
			return '';
		}
		
		function ResetDataChanged()
		{
			datachanged = false;
			return '';
		}

		function HasDataChanged()
		{
			if (datachanged) return 1;
			else return 0;
		}
		
		function GetDataChangedResult()
		{
			return datachanged_result;
		}
		
	]]></msxsl:script>
	<xsl:template match="/">
		<xsl:element name="RESPONSE">
			<xsl:attribute name="TYPE">SUCCESS</xsl:attribute>
			<xsl:for-each select="REQUEST/BEFORE/APPLICATION">
				<xsl:variable name="vAppNumber">
					<xsl:value-of select="@APPLICATIONNUMBER"/>
				</xsl:variable>
				<!-- Credit check test set? -->
				<xsl:if test="//REQUEST/@CREDITCHECK_TEST = 'YES'">
					<xsl:call-template name="CreditCheckTest"/>
				</xsl:if>
				<!-- Cost modelling test flag set? -->
				<xsl:if test="//REQUEST/@COSTMODELLING_TEST = 'YES'">
					<!-- IK 27/04/2006 EP410 only when QUOTATION exists -->
					<xsl:if test="//REQUEST/AFTER/APPLICATION[@ACTIVEQUOTENUMBER]">
						<xsl:call-template name="ApplicationCostModelTest"/>
						<xsl:for-each select="CUSTOMERROLE/CUSTOMERVERSION">
							<xsl:if test="user:HasDataChanged() = 0">
								<xsl:call-template name="CustomerCostModelTest">
									<xsl:with-param name="pAppNumber">
										<xsl:value-of select="$vAppNumber"/>
									</xsl:with-param>
								</xsl:call-template>
							</xsl:if>
						</xsl:for-each>
					</xsl:if>
				</xsl:if>
				<xsl:value-of select="user:ResetDataChanged()"/>
			</xsl:for-each>
			<xsl:element name="CUSTOMERS"/>
		</xsl:element>
	</xsl:template>
	<!--START: Application cost model test-->
	<xsl:template name="ApplicationCostModelTest">
		<xsl:variable name="vAppNumber">
			<xsl:value-of select="@APPLICATIONNUMBER"/>
		</xsl:variable>
		<xsl:variable name="pAppNumber">
			<xsl:value-of select="$vAppNumber"/>
		</xsl:variable>
				<xsl:variable name="vCn">
			<xsl:value-of select="@CUSTOMERNUMBER"/>
		</xsl:variable>
		<xsl:variable name="vCvn">
			<xsl:value-of select="@CUSTOMERVERSIONNUMBER"/>
		</xsl:variable>

		<!-- for consistency - app number is used as a parameter in other templates, referencing  this avoids errors with cut and paste from other templates -->
		<!-- Check to see if purchase price/estimated value has changed-->
		<xsl:variable name="vPurchasePriceOrEstimatedValue_After">
			<xsl:value-of select="//REQUEST/AFTER/APPLICATION[@APPLICATIONNUMBER=$pAppNumber]/@PURCHASEPRICEORESTIMATEDVALUE"/>
		</xsl:variable>
		
		<xsl:variable name="vPurchasePriceOrEstimatedValue_Before"><xsl:value-of select="@PURCHASEPRICEORESTIMATEDVALUE"/></xsl:variable>
		<xsl:if test="$vPurchasePriceOrEstimatedValue_Before != $vPurchasePriceOrEstimatedValue_After">
			<xsl:if test="user:HasDataChanged() = 0">
				<xsl:element name="APPLICATION_COST_MODEL_CHECK">
					<xsl:attribute name="APPLICATION"><xsl:value-of select="$pAppNumber"/></xsl:attribute>
					<xsl:element name="CHANGED">
						<xsl:attribute name="ID">PURCHASEPRICEORESTIMATEDVALUE</xsl:attribute>
						<xsl:attribute name="before"><xsl:value-of select="$vPurchasePriceOrEstimatedValue_Before"/></xsl:attribute>
						<xsl:attribute name="after"><xsl:value-of select="$vPurchasePriceOrEstimatedValue_After"/></xsl:attribute>
					</xsl:element>
					<xsl:value-of select="user:SetDataChanged()"/>
				</xsl:element>
			</xsl:if>
		</xsl:if>	
					
		<xsl:variable name="vValuationType_Before">
			<xsl:value-of select="NEWPROPERTY[@APPLICATIONNUMBER=$pAppNumber]/@VALUATIONTYPE_TEXT"/>
		</xsl:variable>
		<xsl:variable name="vValuationType_After">
			<xsl:value-of select="//REQUEST/AFTER/APPLICATION[@APPLICATIONNUMBER=$pAppNumber]/NEWPROPERTY[@APPLICATIONNUMBER=$pAppNumber]/@VALUATIONTYPE_TEXT"/>
		</xsl:variable>

		<!-- Has valuation type changed from 'No Valuation Required' ? -->
		<xsl:if test="$vValuationType_Before = 'No Valuation Required'">
			<xsl:if test="$vValuationType_After != 'No Valuation Required' ">
				<xsl:if test="user:HasDataChanged() = 0">
					<xsl:element name="APPLICATION_COST_MODEL_CHECK">
						<xsl:attribute name="APPLICATION"><xsl:value-of select="$pAppNumber"/></xsl:attribute>
						<xsl:element name="CHANGED">
							<xsl:attribute name="ID">VALUATIONTYPE</xsl:attribute>
							<xsl:attribute name="before"><xsl:value-of select="$vValuationType_Before"/></xsl:attribute>
							<xsl:attribute name="after"><xsl:value-of select="$vValuationType_After"/></xsl:attribute>
						</xsl:element>
						<xsl:value-of select="user:SetDataChanged()"/>
					</xsl:element>
				</xsl:if>
			</xsl:if>
		</xsl:if>

	<!-- Has valuation type changed to 'No Valuation Required' ? -->
		<xsl:if test="$vValuationType_Before != 'No Valuation Required'">
			<xsl:if test="$vValuationType_After = 'No Valuation Required' ">
				<xsl:if test="user:HasDataChanged() = 0">
					<xsl:element name="APPLICATION_COST_MODEL_CHECK">
						<xsl:attribute name="APPLICATION"><xsl:value-of select="$pAppNumber"/></xsl:attribute>
						<xsl:element name="CHANGED">
							<xsl:attribute name="ID">VALUATIONTYPE</xsl:attribute>
							<xsl:attribute name="before"><xsl:value-of select="$vValuationType_Before"/></xsl:attribute>
							<xsl:attribute name="after"><xsl:value-of select="$vValuationType_After"/></xsl:attribute>
						</xsl:element>
						<xsl:value-of select="user:SetDataChanged()"/>
					</xsl:element>
				</xsl:if>
			</xsl:if>
		</xsl:if>
		
		<xsl:variable name="vRoomType_Bedroom_Before">
			<xsl:value-of select="NEWPROPERTY[@APPLICATIONNUMBER=$pAppNumber]/NEWPROPERTYROOMTYPE/@ROOMTYPE_TYPE_BD"/>
		</xsl:variable>
		<xsl:variable name="vRoomType_Bedroom_After">
			<xsl:value-of select="//REQUEST/AFTER/APPLICATION[@APPLICATIONNUMBER=$pAppNumber]/NEWPROPERTY[@APPLICATIONNUMBER=$pAppNumber]/NEWPROPERTYROOMTYPE/@ROOMTYPE_TYPE_BD"/>
		</xsl:variable>
		<xsl:if test="$vRoomType_Bedroom_Before = 'true' and $vRoomType_Bedroom_After = 'true' ">
			<xsl:variable name="vBedroomCount_Before">
			<xsl:value-of select="NEWPROPERTY[@APPLICATIONNUMBER=$pAppNumber]/NEWPROPERTYROOMTYPE/@NUMBEROFROOMS"/>
			</xsl:variable>
			<xsl:variable name="vBedroomCount_After">
			<xsl:value-of select="//REQUEST/AFTER/APPLICATION[@APPLICATIONNUMBER=$pAppNumber]/NEWPROPERTY[@APPLICATIONNUMBER=$pAppNumber]/NEWPROPERTYROOMTYPE/@NUMBEROFROOMS"/>
			</xsl:variable>
			<xsl:if test="$vBedroomCount_Before != $vBedroomCount_After">
				<xsl:if test="user:HasDataChanged() = 0">
					<xsl:element name="APPLICATION_COST_MODEL_CHECK">
						<xsl:attribute name="APPLICATION"><xsl:value-of select="$pAppNumber"/></xsl:attribute>
						<xsl:element name="CHANGED">
							<xsl:attribute name="ID">ROOMTYPE_TYPE_BD (BEDROOMS)</xsl:attribute>
							<xsl:attribute name="before"><xsl:value-of select="$vBedroomCount_Before"/></xsl:attribute>
							<xsl:attribute name="after"><xsl:value-of select="$vBedroomCount_After"/></xsl:attribute>
						</xsl:element>
						<xsl:value-of select="user:SetDataChanged()"/>
					</xsl:element>
				</xsl:if>
			</xsl:if>
		</xsl:if>
		
		<xsl:variable name="vNatureOfLoan_After">
		<xsl:value-of select="//REQUEST/AFTER/APPLICATION[@APPLICATIONNUMBER=$pAppNumber]/@NATUREOFLOAN"/>
		</xsl:variable>
		<xsl:variable name="vNatureOfLoan_Before"><xsl:value-of select="@NATUREOFLOAN"/></xsl:variable>
		<xsl:if test="$vNatureOfLoan_Before != $vNatureOfLoan_After">
			<xsl:if test="user:HasDataChanged() = 0">
				<xsl:element name="APPLICATION_COST_MODEL_CHECK">
					<xsl:attribute name="APPLICATION"><xsl:value-of select="$pAppNumber"/></xsl:attribute>
					<xsl:element name="CHANGED">
						<xsl:attribute name="ID">NATUREOFLOAN</xsl:attribute>
						<xsl:attribute name="before"><xsl:value-of select="$vNatureOfLoan_Before"/></xsl:attribute>
						<xsl:attribute name="after"><xsl:value-of select="$vNatureOfLoan_After"/></xsl:attribute>
					</xsl:element>
					<xsl:value-of select="user:SetDataChanged()"/>
				</xsl:element>
			</xsl:if>
		</xsl:if>

		<!-- LH 17/05/2006 EP558 - SelfCertInd check-->
		<xsl:variable name="vSelfCertInd_After">
		<xsl:value-of select="//REQUEST/AFTER/APPLICATION[@APPLICATIONNUMBER=$pAppNumber]/@SELFCERTIND"/>
		</xsl:variable>
		<xsl:variable name="vSelfCertInd_Before"><xsl:value-of select="@SELFCERTIND"/></xsl:variable>
		<xsl:if test="$vSelfCertInd_Before != $vSelfCertInd_After">
			<xsl:if test="user:HasDataChanged() = 0">
				<xsl:element name="RESCORE_CREDIT_CHECK">
					<xsl:attribute name="APPLICATION"><xsl:value-of select="$pAppNumber"/></xsl:attribute>
					<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="@CUSTOMERNUMBER"/></xsl:attribute>
					<xsl:element name="CHANGED">
						<xsl:attribute name="ID">SELFCERTIND</xsl:attribute>
						<xsl:attribute name="before"><xsl:value-of select="$vSelfCertInd_Before"/></xsl:attribute>
						<xsl:attribute name="after"><xsl:value-of select="$vSelfCertInd_After"/></xsl:attribute>
					</xsl:element>
					<xsl:value-of select="user:SetDataChanged()"/>
				</xsl:element>
			</xsl:if>
		</xsl:if>

			<!-- SR 11/03/2007: EP2_1828 - check if any of the records in INTRODUCERFEE changed with FeeType validation APR'-->
			<xsl:if test="user:HasDataChanged() = 0">
				<xsl:for-each select="//REQUEST/AFTER/APPLICATION[@APPLICATIONNUMBER=$pAppNumber]/INTRODUCERFEE">
						<xsl:if test="user:HasDataChanged() = 0">
							<xsl:variable name="vFeeType"><xsl:value-of select="@FEETYPE"/></xsl:variable>
							<xsl:variable name="vFeeSeqNo"><xsl:value-of select="@FEESEQUENCENUMBER"/></xsl:variable>
							<xsl:variable name="vFeeAmount"><xsl:value-of select="@FEEAMOUNT"/></xsl:variable>
							<xsl:variable name="vFeeTypeAPR">
							<xsl:choose>
									<xsl:when test="//REQUEST/AFTER/APPLICATION/COMBOGROUP[@GROUPNAME='OneOffCost']/COMBOVALUE[@VALUEID=$vFeeType]/COMBOVALIDATION[@VALIDATIONTYPE='APR']">1</xsl:when>
									<xsl:otherwise>0</xsl:otherwise>
								</xsl:choose>
							</xsl:variable>
							<xsl:if test="$vFeeTypeAPR=1">
								<xsl:call-template name="IntroducerFeeChanged">
									<xsl:with-param name="pFeeType"><xsl:value-of select="$vFeeType"/></xsl:with-param>
									<xsl:with-param name="pFeeSeqNo"><xsl:value-of select="$vFeeSeqNo"/></xsl:with-param>
									<xsl:with-param name="pFeeAmount"><xsl:value-of select="$vFeeAmount"/></xsl:with-param>
									<xsl:with-param name="pApplicationNumber"><xsl:value-of select="$pAppNumber"/></xsl:with-param>
								</xsl:call-template>
							</xsl:if>
						</xsl:if>
				</xsl:for-each>
				<!-- Any of the INTRODUCERFEE records have been deleted (with FeeType validation 'APR')-->
				<xsl:if test="user:HasDataChanged() = 0">
					<xsl:for-each select="//REQUEST/BEFORE/APPLICATION[@APPLICATIONNUMBER=$pAppNumber]/INTRODUCERFEE"> 
						<xsl:variable name="vFeeType"><xsl:value-of select="@FEETYPE"/></xsl:variable>
						<xsl:variable name="vFeeSeqNo"><xsl:value-of select="@FEESEQUENCENUMBER"/></xsl:variable>		
						 <xsl:if test="user:HasDataChanged() = 0">
							<xsl:if test="//REQUEST/AFTER/APPLICATION/COMBOGROUP[@GROUPNAME='OneOffCost']/COMBOVALUE[@VALUEID=$vFeeType]/COMBOVALIDATION[@VALIDATIONTYPE='APR']">
								<xsl:if test="not(//REQUEST/AFTER/APPLICATION[@APPLICATIONNUMBER=$pAppNumber]/INTRODUCERFEE[@FEETYPE=$vFeeType and @FEESEQUENCENUMBER=vFeeSeqNo])">
										<xsl:value-of select="user:SetDataChanged()"/>
									</xsl:if>
							</xsl:if>
						</xsl:if>
					</xsl:for-each>
				</xsl:if>
			</xsl:if>
		<!-- SR 11/03/2007: EP2_1828 - End -->							
		<xsl:if test="user:HasDataChanged() = 1">
			<xsl:call-template name="CreateTask">
				<xsl:with-param name="task">TMRemodelMortgage</xsl:with-param>
			</xsl:call-template>
			<xsl:value-of select="user:ResetDataChanged()"/>
		</xsl:if>
	</xsl:template>
	<!--END: Application cost model test-->
	<!--START: Customer cost model test-->
	<xsl:template name="CustomerCostModelTest">
		<xsl:param name="pAppNumber"/>
		<xsl:variable name="vCn">
			<xsl:value-of select="@CUSTOMERNUMBER"/>
		</xsl:variable>
		<xsl:variable name="vCvn">
			<xsl:value-of select="@CUSTOMERVERSIONNUMBER"/>
		</xsl:variable>
		<xsl:variable name="vGender_After">
			<xsl:value-of select="//REQUEST/AFTER/APPLICATION[@APPLICATIONNUMBER=$pAppNumber]/CUSTOMERROLE/CUSTOMERVERSION[@CUSTOMERNUMBER=$vCn][@CUSTOMERVERSIONNUMBER=$vCvn]/@GENDER"/>
		</xsl:variable>
		<xsl:variable name="vGender_Before"><xsl:value-of select="@GENDER"/></xsl:variable>
		<xsl:if test="$vGender_Before != $vGender_After">
			<xsl:if test="user:HasDataChanged() = 0">
				<xsl:element name="CUSTOMER_COST_MODEL_CHECK">
					<xsl:attribute name="APPLICATION"><xsl:value-of select="$pAppNumber"/></xsl:attribute>
					<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="@CUSTOMERNUMBER"/></xsl:attribute>
					<xsl:element name="CHANGED">
						<xsl:attribute name="ID">GENDER</xsl:attribute>
						<xsl:attribute name="before"><xsl:value-of select="$vGender_Before"/></xsl:attribute>
						<xsl:attribute name="after"><xsl:value-of select="$vGender_After"/></xsl:attribute>
					</xsl:element>
					<xsl:value-of select="user:SetDataChanged()"/>
				</xsl:element>
			</xsl:if>
		</xsl:if>
		<xsl:variable name="vMemberOfStaff_After">
			<xsl:value-of select="//REQUEST/AFTER/APPLICATION[@APPLICATIONNUMBER=$pAppNumber]/CUSTOMERROLE/CUSTOMERVERSION[@CUSTOMERNUMBER=$vCn][@CUSTOMERVERSIONNUMBER=$vCvn]/@MEMBEROFSTAFF"/>
		</xsl:variable>
		<xsl:variable name="vMemberOfStaff_Before"><xsl:value-of select="@MEMBEROFSTAFF"/></xsl:variable>
		<xsl:if test="$vMemberOfStaff_Before != $vMemberOfStaff_After">
			<xsl:if test="user:HasDataChanged() = 0">
				<xsl:element name="CUSTOMER_COST_MODEL_CHECK">
					<xsl:attribute name="APPLICATION"><xsl:value-of select="$pAppNumber"/></xsl:attribute>
					<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="$vCn"/></xsl:attribute>
					<xsl:element name="CHANGED">
						<xsl:attribute name="ID">MEMBER_OF_STAFF</xsl:attribute>
						<xsl:attribute name="before"><xsl:value-of select="$vMemberOfStaff_Before"/></xsl:attribute>
						<xsl:attribute name="after"><xsl:value-of select="$vMemberOfStaff_After"/></xsl:attribute>
					</xsl:element>
					<xsl:value-of select="user:SetDataChanged()"/>
				</xsl:element>
			</xsl:if>
		</xsl:if>
		<xsl:variable name="vNationality_After">
			<xsl:value-of select="//REQUEST/AFTER/APPLICATION[@APPLICATIONNUMBER=$pAppNumber]/CUSTOMERROLE/CUSTOMERVERSION[@CUSTOMERNUMBER=$vCn][@CUSTOMERVERSIONNUMBER=$vCvn]/@NATIONALITY"/>
		</xsl:variable>
		<xsl:variable name="vNationality_Before"><xsl:value-of select="@NATIONALITY"/></xsl:variable>
		<xsl:if test="$vNationality_Before != $vNationality_After">
			<xsl:if test="user:HasDataChanged() = 0">
				<xsl:element name="CUSTOMER_COST_MODEL_CHECK">
					<xsl:attribute name="APPLICATION"><xsl:value-of select="$pAppNumber"/></xsl:attribute>
					<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="@CUSTOMERNUMBER"/></xsl:attribute>
					<xsl:element name="CHANGED">
						<xsl:attribute name="ID">NATIONALITY</xsl:attribute>
						<xsl:attribute name="before"><xsl:value-of select="$vNationality_Before"/></xsl:attribute>
						<xsl:attribute name="after"><xsl:value-of select="$vNationality_After"/></xsl:attribute>
					</xsl:element>
					<xsl:value-of select="user:SetDataChanged()"/>
				</xsl:element>
			</xsl:if>
		</xsl:if>
		<xsl:variable name="vUKResidentInd_After">
			<xsl:value-of select="//REQUEST/AFTER/APPLICATION[@APPLICATIONNUMBER=$pAppNumber]/CUSTOMERROLE/CUSTOMERVERSION[@CUSTOMERNUMBER=$vCn][@CUSTOMERVERSIONNUMBER=$vCvn]/@UKRESIDENTINDICATOR"/>
		</xsl:variable>
		<xsl:variable name="vUKResidentInd_Before"><xsl:value-of select="@UKRESIDENTINDICATOR"/></xsl:variable>
		<xsl:if test="$vUKResidentInd_Before != $vUKResidentInd_After">
			<xsl:if test="user:HasDataChanged() = 0">
				<xsl:element name="CUSTOMER_COST_MODEL_CHECK">
					<xsl:attribute name="APPLICATION"><xsl:value-of select="$pAppNumber"/></xsl:attribute>
					<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="@CUSTOMERNUMBER"/></xsl:attribute>
					<xsl:element name="CHANGED">
						<xsl:attribute name="ID">UK RESIDENT INDICATOR</xsl:attribute>
						<xsl:attribute name="before"><xsl:value-of select="$vUKResidentInd_Before"/></xsl:attribute>
						<xsl:attribute name="after"><xsl:value-of select="$vUKResidentInd_After"/></xsl:attribute>
					</xsl:element>
					<xsl:value-of select="user:SetDataChanged()"/>
				</xsl:element>
			</xsl:if>
		</xsl:if>
		
		<!-- Need to get addressguid from customeraddress table to extract other current address details from address table-->
		<xsl:variable name="vAddressGuid">
			<xsl:value-of select="	//REQUEST/AFTER/APPLICATION[@APPLICATIONNUMBER=$pAppNumber]/CUSTOMERROLE/CUSTOMERADDRESS[@CUSTOMERNUMBER=$vCn][@CUSTOMERVERSIONNUMBER=$vCvn]/@ADDRESSGUID"/>
		</xsl:variable>
		<xsl:variable name="vPostcode_Before">
			<xsl:value-of select="../CUSTOMERADDRESS[@CUSTOMERNUMBER=$vCn][@CUSTOMERVERSIONNUMBER=$vCvn]/ADDRESS[@ADDRESSGUID=$vAddressGuid]/@POSTCODE"/>
		</xsl:variable>
		<xsl:variable name="vPostcode_After">
			<xsl:value-of select="//REQUEST/AFTER/APPLICATION[@APPLICATIONNUMBER=$pAppNumber]/CUSTOMERROLE/CUSTOMERADDRESS[@CUSTOMERNUMBER=$vCn][@CUSTOMERVERSIONNUMBER=$vCvn]/ADDRESS[@ADDRESSGUID=$vAddressGuid]/@POSTCODE"/>
		</xsl:variable>
		<xsl:if test="$vPostcode_Before != $vPostcode_After">
			<xsl:if test="user:HasDataChanged() = 0">
				<xsl:element name="APPLICATION_COST_MODEL_CHECK">
					<xsl:attribute name="APPLICATION"><xsl:value-of select="$pAppNumber"/></xsl:attribute>
					<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="@CUSTOMERNUMBER"/></xsl:attribute>
					<xsl:element name="CHANGED">
						<xsl:attribute name="ID">POSTCODE (CURRENT ADDRESS)</xsl:attribute>
						<xsl:attribute name="before"><xsl:value-of select="$vPostcode_Before"/></xsl:attribute>
						<xsl:attribute name="after"><xsl:value-of select="$vPostcode_After"/></xsl:attribute>
					</xsl:element>
					<xsl:value-of select="user:SetDataChanged()"/>
				</xsl:element>
			</xsl:if>
		</xsl:if>
			  
		<xsl:if test="user:HasDataChanged() = 1">
			<xsl:call-template name="CreateTask">
				<xsl:with-param name="task">TMRemodelMortgage</xsl:with-param>
			</xsl:call-template>
		</xsl:if>
	</xsl:template>
	<!--END: Customer cost model test-->
	<!--START: Credit check test: run both reprocess and rescoring tests-->
	<xsl:template name="CreditCheckTest">
		<xsl:variable name="vAppNumber">
			<xsl:value-of select="@APPLICATIONNUMBER"/>
		</xsl:variable>
		<xsl:for-each select="CUSTOMERROLE/CUSTOMERVERSION">
			<xsl:if test="user:HasDataChanged() = 0">
				<xsl:call-template name="ReProcessCreditCheck">
					<xsl:with-param name="pAppNumber">
						<xsl:value-of select="$vAppNumber"/>
					</xsl:with-param>
				</xsl:call-template>
			</xsl:if>
		</xsl:for-each>
		<xsl:value-of select="user:ResetDataChanged()"/>
		<xsl:for-each select="CUSTOMERROLE/CUSTOMERVERSION">
			<xsl:if test="user:HasDataChanged() = 0">
				<xsl:call-template name="ReScoreCreditCheck">
					<xsl:with-param name="pAppNumber">
						<xsl:value-of select="$vAppNumber"/>
					</xsl:with-param>
				</xsl:call-template>
			</xsl:if>
		</xsl:for-each>
	</xsl:template>
	<!--END: Credit check test-->
	<!--START: Reprocess credit check-->
	<xsl:template name="ReProcessCreditCheck">
		<xsl:param name="pAppNumber"/>
		<xsl:variable name="vCn">
			<xsl:value-of select="@CUSTOMERNUMBER"/>
		</xsl:variable>
		<xsl:variable name="vCvn">
			<xsl:value-of select="@CUSTOMERVERSIONNUMBER"/>
		</xsl:variable>
		<xsl:variable name="vCustomerCount_Before">
			<xsl:value-of select="count(//REQUEST/BEFORE/APPLICATION[@APPLICATIONNUMBER=$pAppNumber]/CUSTOMERROLE)"/>
		</xsl:variable>
		<xsl:variable name="vCustomerCount_After">
			<xsl:value-of select="count(//REQUEST/AFTER/APPLICATION[@APPLICATIONNUMBER=$pAppNumber]/CUSTOMERROLE)"/>
		</xsl:variable>
		<!--for all customer records, check count hasn't changed-->
		<xsl:if test="$vCustomerCount_Before != $vCustomerCount_After">
			<xsl:if test="user:HasDataChanged() = 0">
				<xsl:element name="REPROCESS_CREDIT_CHECK">
					<xsl:attribute name="APPLICATION"><xsl:value-of select="$pAppNumber"/></xsl:attribute>
					<xsl:element name="CHANGED">
						<xsl:attribute name="ID">CUSTOMER ROLE COUNT</xsl:attribute>
						<xsl:attribute name="before"><xsl:value-of select="$vCustomerCount_Before"/></xsl:attribute>
						<xsl:attribute name="after"><xsl:value-of select="$vCustomerCount_After"/></xsl:attribute>
					</xsl:element>
					<xsl:value-of select="user:SetDataChanged()"/>
				</xsl:element>
			</xsl:if>
		</xsl:if>
		<xsl:variable name="vDOB_After">
		<xsl:value-of select="//REQUEST/AFTER/APPLICATION[@APPLICATIONNUMBER=$pAppNumber]/CUSTOMERROLE/CUSTOMERVERSION[@CUSTOMERNUMBER=$vCn][@CUSTOMERVERSIONNUMBER=$vCvn]/@DATEOFBIRTH"/>
		</xsl:variable>
		<xsl:variable name="vDOB_Before"><xsl:value-of select="@DATEOFBIRTH"/></xsl:variable>
		<xsl:if test="$vDOB_Before != $vDOB_After">
			<xsl:if test="user:HasDataChanged() = 0">
				<xsl:element name="REPROCESS_CREDIT_CHECK">
					<xsl:attribute name="APPLICATION"><xsl:value-of select="$pAppNumber"/></xsl:attribute>
					<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="@CUSTOMERNUMBER"/></xsl:attribute>
					<xsl:element name="CHANGED">
						<xsl:attribute name="ID">DATE OF BIRTH</xsl:attribute>
						<xsl:attribute name="before"><xsl:value-of select="$vDOB_Before"/></xsl:attribute>
						<xsl:attribute name="after"><xsl:value-of select="$vDOB_After"/></xsl:attribute>
					</xsl:element>
					<xsl:value-of select="user:SetDataChanged()"/>
				</xsl:element>
			</xsl:if>
		</xsl:if>
		<xsl:variable name="vFirstforename_After">
		<xsl:value-of select="//REQUEST/AFTER/APPLICATION[@APPLICATIONNUMBER=$pAppNumber]/CUSTOMERROLE/CUSTOMERVERSION[@CUSTOMERNUMBER=$vCn][@CUSTOMERVERSIONNUMBER=$vCvn]/@FIRSTFORENAME"/>
		</xsl:variable>
		<xsl:variable name="vFirstforename_Before"><xsl:value-of select="@FIRSTFORENAME"/></xsl:variable>
		<xsl:if test="$vFirstforename_Before != $vFirstforename_After">
			<xsl:if test="user:HasDataChanged() = 0">
				<xsl:element name="REPROCESS_CREDIT_CHECK">
					<xsl:attribute name="APPLICATION"><xsl:value-of select="$pAppNumber"/></xsl:attribute>
					<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="@CUSTOMERNUMBER"/></xsl:attribute>
					<xsl:element name="CHANGED">
						<xsl:attribute name="ID">FIRST FORENAME</xsl:attribute>
						<xsl:attribute name="before"><xsl:value-of select="$vFirstforename_Before"/></xsl:attribute>
						<xsl:attribute name="after"><xsl:value-of select="$vFirstforename_After"/></xsl:attribute>
					</xsl:element>
				</xsl:element>
				<xsl:value-of select="user:SetDataChanged()"/>
			</xsl:if>
		</xsl:if>
		<xsl:variable name="vSecondforename_After">
			<xsl:value-of select="//REQUEST/AFTER/APPLICATION[@APPLICATIONNUMBER=$pAppNumber]/CUSTOMERROLE/CUSTOMERVERSION[@CUSTOMERNUMBER=$vCn][@CUSTOMERVERSIONNUMBER=$vCvn]/@SECONDFORENAME"/>
		</xsl:variable>
		<xsl:variable name="vSecondforename_Before"><xsl:value-of select="@SECONDFORENAME"/></xsl:variable>
		<xsl:if test="$vSecondforename_Before != $vSecondforename_After">
			<xsl:if test="user:HasDataChanged() = 0">
				<xsl:element name="REPROCESS_CREDIT_CHECK">
					<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="@CUSTOMERNUMBER"/></xsl:attribute>
					<xsl:element name="CHANGED">
						<xsl:attribute name="ID">SECOND FORENAME</xsl:attribute>
						<xsl:attribute name="before"><xsl:value-of select="$vSecondforename_Before"/></xsl:attribute>
						<xsl:attribute name="after"><xsl:value-of select="$vSecondforename_After"/></xsl:attribute>
					</xsl:element>
				</xsl:element>
				<xsl:value-of select="user:SetDataChanged()"/>
			</xsl:if>
		</xsl:if>
		<xsl:variable name="vSurname_After">
		<xsl:value-of select="//REQUEST/AFTER/APPLICATION[@APPLICATIONNUMBER=$pAppNumber]/CUSTOMERROLE/CUSTOMERVERSION[@CUSTOMERNUMBER=$vCn][@CUSTOMERVERSIONNUMBER=$vCvn]/@SURNAME"/>
		</xsl:variable>
		<xsl:variable name="vSurname_Before"><xsl:value-of select="@SURNAME"/></xsl:variable>
		<xsl:if test="$vSurname_Before != $vSurname_After">
			<xsl:if test="user:HasDataChanged() = 0">
				<xsl:element name="REPROCESS_CREDIT_CHECK">
					<xsl:attribute name="APPLICATION"><xsl:value-of select="$pAppNumber"/></xsl:attribute>
					<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="@CUSTOMERNUMBER"/></xsl:attribute>
					<xsl:element name="CHANGED">
						<xsl:attribute name="ID">SURNAME</xsl:attribute>
						<xsl:attribute name="before"><xsl:value-of select="$vSurname_Before"/></xsl:attribute>
						<xsl:attribute name="after"><xsl:value-of select="$vSurname_After"/></xsl:attribute>
					</xsl:element>
				</xsl:element>
				<xsl:value-of select="user:SetDataChanged()"/>
			</xsl:if>
		</xsl:if>
		<xsl:variable name="vTitle_After">
		<xsl:value-of select="//REQUEST/AFTER/APPLICATION[@APPLICATIONNUMBER=$pAppNumber]/CUSTOMERROLE/CUSTOMERVERSION[@CUSTOMERNUMBER=$vCn][@CUSTOMERVERSIONNUMBER=$vCvn]/@TITLE"/>
		</xsl:variable>
		<xsl:variable name="vTitle_Before"><xsl:value-of select="@TITLE"/></xsl:variable>
		<xsl:if test="$vTitle_Before != $vTitle_After">
			<xsl:if test="user:HasDataChanged() = 0">
				<xsl:element name="REPROCESS_CREDIT_CHECK">
					<xsl:attribute name="APPLICATION"><xsl:value-of select="$pAppNumber"/></xsl:attribute>
					<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="@CUSTOMERNUMBER"/></xsl:attribute>
					<xsl:element name="CHANGED">
						<xsl:attribute name="ID">TITLE</xsl:attribute>
						<xsl:attribute name="before"><xsl:value-of select="$vTitle_Before"/></xsl:attribute>
						<xsl:attribute name="after"><xsl:value-of select="$vTitle_After"/></xsl:attribute>
					</xsl:element>
				</xsl:element>
				<xsl:value-of select="user:SetDataChanged()"/>
			</xsl:if>
		</xsl:if>
		<xsl:variable name="vGender_After">
		<xsl:value-of select="//REQUEST/AFTER/APPLICATION[@APPLICATIONNUMBER=$pAppNumber]/CUSTOMERROLE/CUSTOMERVERSION[@CUSTOMERNUMBER=$vCn][@CUSTOMERVERSIONNUMBER=$vCvn]/@GENDER"/>
		</xsl:variable>
		<xsl:variable name="vGender_Before"><xsl:value-of select="@GENDER"/></xsl:variable>
		<xsl:if test="$vGender_Before != $vGender_After">
			<xsl:if test="user:HasDataChanged() = 0">
				<xsl:element name="REPROCESS_CREDIT_CHECK">
					<xsl:attribute name="APPLICATION"><xsl:value-of select="$pAppNumber"/></xsl:attribute>
					<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="@CUSTOMERNUMBER"/></xsl:attribute>
					<xsl:element name="CHANGED">
						<xsl:attribute name="ID">GENDER</xsl:attribute>
						<xsl:attribute name="before"><xsl:value-of select="$vGender_Before"/></xsl:attribute>
						<xsl:attribute name="after"><xsl:value-of select="$vGender_After"/></xsl:attribute>
					</xsl:element>
					<xsl:value-of select="user:SetDataChanged()"/>
				</xsl:element>
			</xsl:if>
		</xsl:if>
		<!-- Search through customer address' where home or previous-->
		<xsl:for-each select="../CUSTOMERADDRESS[@ADDRESSTYPE_TYPE_H='true' or @ADDRESSTYPE_TYPE_P='true']">
			<xsl:variable name="vAddressSequenceNumber"><xsl:value-of select="@CUSTOMERADDRESSSEQUENCENUMBER"/></xsl:variable>
			<xsl:variable name="vDateMovedIn_After">
				<xsl:value-of select="//REQUEST/AFTER/APPLICATION[@APPLICATIONNUMBER=$pAppNumber]/CUSTOMERROLE/CUSTOMERADDRESS[@CUSTOMERNUMBER=$vCn][@CUSTOMERVERSIONNUMBER=$vCvn][@CUSTOMERADDRESSSEQUENCENUMBER=$vAddressSequenceNumber]/@DATEMOVEDIN"/>
			</xsl:variable>
			<xsl:variable name="vDateMovedIn_Before"><xsl:value-of select="@DATEMOVEDIN"/></xsl:variable>
			<xsl:if test="$vDateMovedIn_Before != $vDateMovedIn_After">
				<xsl:if test="user:HasDataChanged() = 0">
					<xsl:element name="REPROCESS_CREDIT_CHECK">
						<xsl:attribute name="APPLICATION"><xsl:value-of select="$pAppNumber"/></xsl:attribute>
						<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="@CUSTOMERNUMBER"/></xsl:attribute>
						<xsl:element name="CHANGED">
							<xsl:attribute name="ID">DATEMOVEDIN</xsl:attribute>
							<xsl:attribute name="before"><xsl:value-of select="$vDateMovedIn_Before"/></xsl:attribute>
							<xsl:attribute name="after"><xsl:value-of select="$vDateMovedIn_After"/></xsl:attribute>
						</xsl:element>
						<xsl:value-of select="user:SetDataChanged()"/>
					</xsl:element>
				</xsl:if>
			</xsl:if>
			
			<!--Previous address only-->
			<xsl:variable name="vDateMovedOut_After">
			<xsl:value-of select="//REQUEST/AFTER/APPLICATION[@APPLICATIONNUMBER=$pAppNumber]/CUSTOMERROLE/CUSTOMERADDRESS[@CUSTOMERNUMBER=$vCn][@CUSTOMERVERSIONNUMBER=$vCvn][@CUSTOMERADDRESSSEQUENCENUMBER=$vAddressSequenceNumber]/@DATEMOVEDOUT"/>
			</xsl:variable>
			<xsl:variable name="vDateMovedOut_Before"><xsl:value-of select="@DATEMOVEDOUT"/></xsl:variable>
			<xsl:if test="$vDateMovedOut_Before != $vDateMovedOut_After">
				<xsl:if test="user:HasDataChanged() = 0">
					<xsl:element name="REPROCESS_CREDIT_CHECK">
						<xsl:attribute name="APPLICATION"><xsl:value-of select="$pAppNumber"/></xsl:attribute>
						<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="@CUSTOMERNUMBER"/></xsl:attribute>
						<xsl:element name="CHANGED">
							<xsl:attribute name="ID">DATEMOVEDOUT</xsl:attribute>
							<xsl:attribute name="before"><xsl:value-of select="$vDateMovedOut_Before"/></xsl:attribute>
							<xsl:attribute name="after"><xsl:value-of select="$vDateMovedOut_After"/></xsl:attribute>
						</xsl:element>
						<xsl:value-of select="user:SetDataChanged()"/>
					</xsl:element>
				</xsl:if>
			</xsl:if>
			
			<!-- Need to get addressguid from customeraddress table to extract other address details from address table-->
			<xsl:variable name="vAddressGuid">
				<xsl:value-of select="	//REQUEST/AFTER/APPLICATION[@APPLICATIONNUMBER=$pAppNumber]/CUSTOMERROLE/CUSTOMERADDRESS[@CUSTOMERNUMBER=$vCn][@CUSTOMERVERSIONNUMBER=$vCvn][@CUSTOMERADDRESSSEQUENCENUMBER=$vAddressSequenceNumber]/@ADDRESSGUID"/>
			</xsl:variable>
			<xsl:variable name="vBuildingorHouseName_Before">
				<xsl:value-of select="ADDRESS[@ADDRESSGUID=$vAddressGuid]/@BUILDINGORHOUSENAME"/>
			</xsl:variable>
			<xsl:variable name="vBuildingorHouseName_After">
				<xsl:value-of select="//REQUEST/AFTER/APPLICATION[@APPLICATIONNUMBER=$pAppNumber]/CUSTOMERROLE/CUSTOMERADDRESS[@CUSTOMERNUMBER=$vCn][@CUSTOMERVERSIONNUMBER=$vCvn]/ADDRESS[@ADDRESSGUID=$vAddressGuid]/@BUILDINGORHOUSENAME"/>
			</xsl:variable>
			<xsl:if test="$vBuildingorHouseName_Before != $vBuildingorHouseName_After">
				<xsl:if test="user:HasDataChanged() = 0">
					<xsl:element name="REPROCESS_CREDIT_CHECK">
						<xsl:attribute name="APPLICATION"><xsl:value-of select="$pAppNumber"/></xsl:attribute>
						<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="@CUSTOMERNUMBER"/></xsl:attribute>
						<xsl:element name="CHANGED">
							<xsl:attribute name="ID">BUILDINGORHOUSENAME</xsl:attribute>
							<xsl:attribute name="before"><xsl:value-of select="$vBuildingorHouseName_Before"/></xsl:attribute>
							<xsl:attribute name="after"><xsl:value-of select="$vBuildingorHouseName_After"/></xsl:attribute>
						</xsl:element>
						<xsl:value-of select="user:SetDataChanged()"/>
					</xsl:element>
				</xsl:if>
			</xsl:if>
			<xsl:variable name="vBuildingorHouseNo_Before">
				<xsl:value-of select="ADDRESS[@ADDRESSGUID=$vAddressGuid]/@BUILDINGORHOUSENUMBER"/>
			</xsl:variable>
			<xsl:variable name="vBuildingorHouseNo_After">
				<xsl:value-of select="//REQUEST/AFTER/APPLICATION[@APPLICATIONNUMBER=$pAppNumber]/CUSTOMERROLE/CUSTOMERADDRESS[@CUSTOMERNUMBER=$vCn][@CUSTOMERVERSIONNUMBER=$vCvn]/ADDRESS[@ADDRESSGUID=$vAddressGuid]/@BUILDINGORHOUSENUMBER"/>
			</xsl:variable>
			<xsl:if test="$vBuildingorHouseNo_Before != $vBuildingorHouseNo_After">
				<xsl:if test="user:HasDataChanged() = 0">
					<xsl:element name="REPROCESS_CREDIT_CHECK">
						<xsl:attribute name="APPLICATION"><xsl:value-of select="$pAppNumber"/></xsl:attribute>
						<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="@CUSTOMERNUMBER"/></xsl:attribute>
						<xsl:element name="CHANGED">
							<xsl:attribute name="ID">BUILDINGORHOUSENUMBER</xsl:attribute>
							<xsl:attribute name="before"><xsl:value-of select="$vBuildingorHouseNo_Before"/></xsl:attribute>
							<xsl:attribute name="after"><xsl:value-of select="$vBuildingorHouseNo_After"/></xsl:attribute>
						</xsl:element>
						<xsl:value-of select="user:SetDataChanged()"/>
					</xsl:element>
				</xsl:if>
			</xsl:if>
			<xsl:variable name="vFlatNo_Before">
				<xsl:value-of select="ADDRESS[@ADDRESSGUID=$vAddressGuid]/@FLATNUMBER"/>
			</xsl:variable>
			<xsl:variable name="vFlatNo_After">
				<xsl:value-of select="//REQUEST/AFTER/APPLICATION[@APPLICATIONNUMBER=$pAppNumber]/CUSTOMERROLE/CUSTOMERADDRESS[@CUSTOMERNUMBER=$vCn][@CUSTOMERVERSIONNUMBER=$vCvn]/ADDRESS[@ADDRESSGUID=$vAddressGuid]/@FLATNUMBER"/>
			</xsl:variable>
			<xsl:if test="$vFlatNo_Before != $vFlatNo_After">
				<xsl:if test="user:HasDataChanged() = 0">
					<xsl:element name="REPROCESS_CREDIT_CHECK">
						<xsl:attribute name="APPLICATION"><xsl:value-of select="$pAppNumber"/></xsl:attribute>
						<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="@CUSTOMERNUMBER"/></xsl:attribute>
						<xsl:element name="CHANGED">
							<xsl:attribute name="ID">FLATNUMBER</xsl:attribute>
							<xsl:attribute name="before"><xsl:value-of select="$vFlatNo_Before"/></xsl:attribute>
							<xsl:attribute name="after"><xsl:value-of select="$vFlatNo_After"/></xsl:attribute>
						</xsl:element>
						<xsl:value-of select="user:SetDataChanged()"/>
					</xsl:element>
				</xsl:if>
			</xsl:if>
			<xsl:variable name="vStreet_Before">
				<xsl:value-of select="ADDRESS[@ADDRESSGUID=$vAddressGuid]/@STREET"/>
			</xsl:variable>
			<xsl:variable name="vStreet_After">
				<xsl:value-of select="//REQUEST/AFTER/APPLICATION[@APPLICATIONNUMBER=$pAppNumber]/CUSTOMERROLE/CUSTOMERADDRESS[@CUSTOMERNUMBER=$vCn][@CUSTOMERVERSIONNUMBER=$vCvn]/ADDRESS[@ADDRESSGUID=$vAddressGuid]/@STREET"/>
			</xsl:variable>
			<xsl:if test="$vStreet_Before != $vStreet_After">
				<xsl:if test="user:HasDataChanged() = 0">
					<xsl:element name="REPROCESS_CREDIT_CHECK">
						<xsl:attribute name="APPLICATION"><xsl:value-of select="$pAppNumber"/></xsl:attribute>
						<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="@CUSTOMERNUMBER"/></xsl:attribute>
						<xsl:element name="CHANGED">
							<xsl:attribute name="ID">STREET</xsl:attribute>
							<xsl:attribute name="before"><xsl:value-of select="$vStreet_Before"/></xsl:attribute>
							<xsl:attribute name="after"><xsl:value-of select="$vStreet_After"/></xsl:attribute>
						</xsl:element>
						<xsl:value-of select="user:SetDataChanged()"/>
					</xsl:element>
				</xsl:if>
			</xsl:if>
			<xsl:variable name="vDistrict_Before">
				<xsl:value-of select="ADDRESS[@ADDRESSGUID=$vAddressGuid]/@DISTRICT"/>
			</xsl:variable>
			<xsl:variable name="vDistrict_After">
				<xsl:value-of select="//REQUEST/AFTER/APPLICATION[@APPLICATIONNUMBER=$pAppNumber]/CUSTOMERROLE/CUSTOMERADDRESS[@CUSTOMERNUMBER=$vCn][@CUSTOMERVERSIONNUMBER=$vCvn]/ADDRESS[@ADDRESSGUID=$vAddressGuid]/@DISTRICT"/>
			</xsl:variable>
			<xsl:if test="$vDistrict_Before != $vDistrict_After">
				<xsl:if test="user:HasDataChanged() = 0">
					<xsl:element name="REPROCESS_CREDIT_CHECK">
						<xsl:attribute name="APPLICATION"><xsl:value-of select="$pAppNumber"/></xsl:attribute>
						<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="@CUSTOMERNUMBER"/></xsl:attribute>
						<xsl:element name="CHANGED">
							<xsl:attribute name="ID">DISTRICT</xsl:attribute>
							<xsl:attribute name="before"><xsl:value-of select="$vDistrict_Before"/></xsl:attribute>
							<xsl:attribute name="after"><xsl:value-of select="$vDistrict_After"/></xsl:attribute>
						</xsl:element>
						<xsl:value-of select="user:SetDataChanged()"/>
					</xsl:element>
				</xsl:if>
			</xsl:if>
			<xsl:variable name="vTown_Before">
				<xsl:value-of select="ADDRESS[@ADDRESSGUID=$vAddressGuid]/@TOWN"/>
			</xsl:variable>
			<xsl:variable name="vTown_After">
				<xsl:value-of select="//REQUEST/AFTER/APPLICATION[@APPLICATIONNUMBER=$pAppNumber]/CUSTOMERROLE/CUSTOMERADDRESS[@CUSTOMERNUMBER=$vCn][@CUSTOMERVERSIONNUMBER=$vCvn]/ADDRESS[@ADDRESSGUID=$vAddressGuid]/@TOWN"/>
			</xsl:variable>
			<xsl:if test="$vTown_Before != $vTown_After">
				<xsl:if test="user:HasDataChanged() = 0">
					<xsl:element name="REPROCESS_CREDIT_CHECK">
						<xsl:attribute name="APPLICATION"><xsl:value-of select="$pAppNumber"/></xsl:attribute>
						<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="@CUSTOMERNUMBER"/></xsl:attribute>
						<xsl:element name="CHANGED">
							<xsl:attribute name="ID">TOWN</xsl:attribute>
							<xsl:attribute name="before"><xsl:value-of select="$vTown_Before"/></xsl:attribute>
							<xsl:attribute name="after"><xsl:value-of select="$vTown_After"/></xsl:attribute>
						</xsl:element>
						<xsl:value-of select="user:SetDataChanged()"/>
					</xsl:element>
				</xsl:if>
			</xsl:if>
			<xsl:variable name="vCountry_Before">
				<xsl:value-of select="ADDRESS[@ADDRESSGUID=$vAddressGuid]/@COUNTRY"/>
			</xsl:variable>
			<xsl:variable name="vCountry_After">
				<xsl:value-of select="//REQUEST/AFTER/APPLICATION[@APPLICATIONNUMBER=$pAppNumber]/CUSTOMERROLE/CUSTOMERADDRESS[@CUSTOMERNUMBER=$vCn][@CUSTOMERVERSIONNUMBER=$vCvn]/ADDRESS[@ADDRESSGUID=$vAddressGuid]/@COUNTRY"/>
			</xsl:variable>
			<xsl:if test="$vCountry_Before != $vCountry_After">
				<xsl:if test="user:HasDataChanged() = 0">
					<xsl:element name="REPROCESS_CREDIT_CHECK">
						<xsl:attribute name="APPLICATION"><xsl:value-of select="$pAppNumber"/></xsl:attribute>
						<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="@CUSTOMERNUMBER"/></xsl:attribute>
						<xsl:element name="CHANGED">
							<xsl:attribute name="ID">COUNTRY</xsl:attribute>
							<xsl:attribute name="before"><xsl:value-of select="$vCountry_Before"/></xsl:attribute>
							<xsl:attribute name="after"><xsl:value-of select="$vCountry_After"/></xsl:attribute>
						</xsl:element>
						<xsl:value-of select="user:SetDataChanged()"/>
					</xsl:element>
				</xsl:if>
			</xsl:if>
			<xsl:variable name="vPostcode_Before">
				<xsl:value-of select="ADDRESS[@ADDRESSGUID=$vAddressGuid]/@POSTCODE"/>
			</xsl:variable>
			<xsl:variable name="vPostcode_After">
				<xsl:value-of select="//REQUEST/AFTER/APPLICATION[@APPLICATIONNUMBER=$pAppNumber]/CUSTOMERROLE/CUSTOMERADDRESS[@CUSTOMERNUMBER=$vCn][@CUSTOMERVERSIONNUMBER=$vCvn]/ADDRESS[@ADDRESSGUID=$vAddressGuid]/@POSTCODE"/>
			</xsl:variable>
			<xsl:if test="$vPostcode_Before != $vPostcode_After">
				<xsl:if test="user:HasDataChanged() = 0">
					<xsl:element name="REPROCESS_CREDIT_CHECK">
						<xsl:attribute name="APPLICATION"><xsl:value-of select="$pAppNumber"/></xsl:attribute>
						<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="@CUSTOMERNUMBER"/></xsl:attribute>
						<xsl:element name="CHANGED">
							<xsl:attribute name="ID">POSTCODE</xsl:attribute>
							<xsl:attribute name="before"><xsl:value-of select="$vPostcode_Before"/></xsl:attribute>
							<xsl:attribute name="after"><xsl:value-of select="$vPostcode_After"/></xsl:attribute>
						</xsl:element>
						<xsl:value-of select="user:SetDataChanged()"/>
					</xsl:element>
				</xsl:if>
			</xsl:if>
		</xsl:for-each>

		<!--applicant alias'-->
		<!-- IK 23/06/2006 EP767 new ALIAS: start -->
		<xsl:choose>
			<xsl:when test="../ALIAS[@ALIASTYPE_TYPE_AL]"><!-- IK 23/06/2006 EP767 new ALIAS: end -->
				<xsl:variable name="vAliasSurname_Before">
					<xsl:value-of select="../ALIAS[@CUSTOMERVERSIONNUMBER=$vCvn]/PERSON/@SURNAME"/>
				</xsl:variable>
				<xsl:variable name="vAliasSurname_After">
					<xsl:value-of select="//REQUEST/AFTER/APPLICATION[@APPLICATIONNUMBER=$pAppNumber]/CUSTOMERROLE/ALIAS[@CUSTOMERVERSIONNUMBER=$vCvn]/PERSON/@SURNAME"/>
				</xsl:variable>
				<xsl:if test="$vAliasSurname_Before != $vAliasSurname_After">
					<xsl:if test="user:HasDataChanged() = 0">
						<xsl:element name="REPROCESS_CREDIT_CHECK">
							<xsl:attribute name="APPLICATION"><xsl:value-of select="$pAppNumber"/></xsl:attribute>
							<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="@CUSTOMERNUMBER"/></xsl:attribute>
							<xsl:element name="CHANGED">
								<xsl:attribute name="ID">ALIAS - SURNAME</xsl:attribute>
								<xsl:attribute name="before"><xsl:value-of select="$vAliasSurname_Before"/></xsl:attribute>
								<xsl:attribute name="after"><xsl:value-of select="$vAliasSurname_After"/></xsl:attribute>
							</xsl:element>
							<xsl:value-of select="user:SetDataChanged()"/>
						</xsl:element>
					</xsl:if>
				</xsl:if>
				
				<xsl:variable name="vAliasFirstForename_Before">
					<xsl:value-of select="../ALIAS[@CUSTOMERVERSIONNUMBER=$vCvn]/PERSON/@FIRSTFORENAME"/>
				</xsl:variable>
				<xsl:variable name="vAliasFirstForename_After">
					<xsl:value-of select="//REQUEST/AFTER/APPLICATION[@APPLICATIONNUMBER=$pAppNumber]/CUSTOMERROLE/ALIAS[@CUSTOMERVERSIONNUMBER=$vCvn]/PERSON/@FIRSTFORENAME"/>
				</xsl:variable>
				<xsl:if test="$vAliasFirstForename_Before != $vAliasFirstForename_After">
					<xsl:if test="user:HasDataChanged() = 0">
						<xsl:element name="REPROCESS_CREDIT_CHECK">
							<xsl:attribute name="APPLICATION"><xsl:value-of select="$pAppNumber"/></xsl:attribute>
							<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="@CUSTOMERNUMBER"/></xsl:attribute>
							<xsl:element name="CHANGED">
								<xsl:attribute name="ID">ALIAS - FIRSTFORENAME</xsl:attribute>
								<xsl:attribute name="before"><xsl:value-of select="$vAliasFirstForename_Before"/></xsl:attribute>
								<xsl:attribute name="after"><xsl:value-of select="$vAliasFirstForename_After"/></xsl:attribute>
							</xsl:element>
							<xsl:value-of select="user:SetDataChanged()"/>
						</xsl:element>
					</xsl:if>
				</xsl:if>
				
				<xsl:variable name="vAliasSecondForename_Before">
					<xsl:value-of select="../ALIAS[@CUSTOMERVERSIONNUMBER=$vCvn]/PERSON/@SECONDFORENAME"/>
				</xsl:variable>
				<xsl:variable name="vAliasSecondForename_After">
					<xsl:value-of select="//REQUEST/AFTER/APPLICATION[@APPLICATIONNUMBER=$pAppNumber]/CUSTOMERROLE/ALIAS[@CUSTOMERVERSIONNUMBER=$vCvn]/PERSON/@SECONDFORENAME"/>
				</xsl:variable>
				<xsl:if test="$vAliasSecondForename_Before != $vAliasSecondForename_After">
					<xsl:if test="user:HasDataChanged() = 0">
						<xsl:element name="REPROCESS_CREDIT_CHECK">
							<xsl:attribute name="APPLICATION"><xsl:value-of select="$pAppNumber"/></xsl:attribute>
							<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="@CUSTOMERNUMBER"/></xsl:attribute>
							<xsl:element name="CHANGED">
								<xsl:attribute name="ID">ALIAS - SECONDFORENAME</xsl:attribute>
								<xsl:attribute name="before"><xsl:value-of select="$vAliasSecondForename_Before"/></xsl:attribute>
								<xsl:attribute name="after"><xsl:value-of select="$vAliasSecondForename_After"/></xsl:attribute>
							</xsl:element>
							<xsl:value-of select="user:SetDataChanged()"/>
						</xsl:element>
					</xsl:if>
				</xsl:if>
				
				<xsl:variable name="vAliasTitleOther_Before">
					<xsl:value-of select="../ALIAS[@CUSTOMERVERSIONNUMBER=$vCvn]/PERSON/@TITLEOTHER"/>
				</xsl:variable>
				<xsl:if test="$vAliasTitleOther_Before != ''">
					<xsl:variable name="vAliasTitleOther_After">
						<xsl:value-of select="//REQUEST/AFTER/APPLICATION[@APPLICATIONNUMBER=$pAppNumber]/CUSTOMERROLE/ALIAS[@CUSTOMERVERSIONNUMBER=$vCvn]/PERSON/@TITLEOTHER"/>
					</xsl:variable>
					<xsl:if test="$vAliasTitleOther_Before != $vAliasTitleOther_After">
						<xsl:if test="user:HasDataChanged() = 0">
							<xsl:element name="REPROCESS_CREDIT_CHECK">
								<xsl:attribute name="APPLICATION"><xsl:value-of select="$pAppNumber"/></xsl:attribute>
								<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="@CUSTOMERNUMBER"/></xsl:attribute>
								<xsl:element name="CHANGED">
									<xsl:attribute name="ID">ALIAS - TITLEOTHER</xsl:attribute>
									<xsl:attribute name="before"><xsl:value-of select="$vAliasTitleOther_Before"/></xsl:attribute>
									<xsl:attribute name="after"><xsl:value-of select="$vAliasTitleOther_After"/></xsl:attribute>
								</xsl:element>
								<xsl:value-of select="user:SetDataChanged()"/>
							</xsl:element>
						</xsl:if>
					</xsl:if>
				</xsl:if>
				
				<xsl:if test="$vAliasTitleOther_Before = ''">
					<xsl:variable name="vAliasTitle_Before">
						<xsl:value-of select="../ALIAS[@CUSTOMERVERSIONNUMBER=$vCvn]/PERSON/@TITLE"/>
					</xsl:variable>
					<xsl:variable name="vAliasTitle_After">
						<xsl:value-of select="//REQUEST/AFTER/APPLICATION[@APPLICATIONNUMBER=$pAppNumber]/CUSTOMERROLE/ALIAS[@CUSTOMERVERSIONNUMBER=$vCvn]/PERSON/@TITLE"/>
					</xsl:variable>
					<xsl:if test="$vAliasTitle_Before != $vAliasTitle_After">
						<xsl:if test="user:HasDataChanged() = 0">
							<xsl:element name="REPROCESS_CREDIT_CHECK">
								<xsl:attribute name="APPLICATION"><xsl:value-of select="$pAppNumber"/></xsl:attribute>
								<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="@CUSTOMERNUMBER"/></xsl:attribute>
								<xsl:element name="CHANGED">
									<xsl:attribute name="ID">ALIAS - TITLE</xsl:attribute>
									<xsl:attribute name="before"><xsl:value-of select="$vAliasTitle_Before"/></xsl:attribute>
									<xsl:attribute name="after"><xsl:value-of select="$vAliasTitle_After"/></xsl:attribute>
								</xsl:element>
								<xsl:value-of select="user:SetDataChanged()"/>
							</xsl:element>
						</xsl:if>
					</xsl:if>
				</xsl:if>
				
				<xsl:variable name="vAliasDOB_Before">
				<xsl:value-of select="../CUSTOMERVERSION[@CUSTOMERVERSIONNUMBER=$vCvn]/@DATEOFBIRTH"/>
				</xsl:variable>
				<xsl:variable name="vAliasDOB_After">
					<xsl:value-of select="//REQUEST/AFTER/APPLICATION[@APPLICATIONNUMBER=$pAppNumber]/CUSTOMERROLE/CUSTOMERVERSION[@CUSTOMERVERSIONNUMBER=$vCvn]/@DATEOFBIRTH"/>
				</xsl:variable>
				<xsl:if test="$vAliasDOB_Before != $vAliasDOB_After">
					<xsl:if test="user:HasDataChanged() = 0">
						<xsl:element name="REPROCESS_CREDIT_CHECK">
							<xsl:attribute name="APPLICATION"><xsl:value-of select="$pAppNumber"/></xsl:attribute>
							<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="@CUSTOMERNUMBER"/></xsl:attribute>
							<xsl:element name="CHANGED">
								<xsl:attribute name="ID">ALIAS - DATEOFBIRTH</xsl:attribute>
								<xsl:attribute name="before"><xsl:value-of select="$vAliasDOB_Before"/></xsl:attribute>
								<xsl:attribute name="after"><xsl:value-of select="$vAliasDOB_After"/></xsl:attribute>
							</xsl:element>
							<xsl:value-of select="user:SetDataChanged()"/>
						</xsl:element>
					</xsl:if>
				</xsl:if>
	
				<xsl:variable name="vAliasGender_Before">
					<xsl:value-of select="../ALIAS[@CUSTOMERVERSIONNUMBER=$vCvn]/PERSON/@GENDER"/>
				</xsl:variable>
				<xsl:variable name="vAliasGender_After">
					<xsl:value-of select="//REQUEST/AFTER/APPLICATION[@APPLICATIONNUMBER=$pAppNumber]/CUSTOMERROLE/ALIAS[@CUSTOMERVERSIONNUMBER=$vCvn]/PERSON/@GENDER"/>
				</xsl:variable>
				<xsl:if test="$vAliasGender_Before != $vAliasGender_After">
					<xsl:if test="user:HasDataChanged() = 0">
						<xsl:element name="REPROCESS_CREDIT_CHECK">
							<xsl:attribute name="APPLICATION"><xsl:value-of select="$pAppNumber"/></xsl:attribute>
							<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="@CUSTOMERNUMBER"/></xsl:attribute>
							<xsl:element name="CHANGED">
								<xsl:attribute name="ID">ALIAS - GENDER</xsl:attribute>
								<xsl:attribute name="before"><xsl:value-of select="$vAliasGender_Before"/></xsl:attribute>
								<xsl:attribute name="after"><xsl:value-of select="$vAliasGender_After"/></xsl:attribute>
							</xsl:element>
							<xsl:value-of select="user:SetDataChanged()"/>
						</xsl:element>
					</xsl:if>
				</xsl:if>
			<!-- IK 23/06/2006 EP767 new ALIAS: start -->
			</xsl:when>
			<xsl:otherwise>
				<xsl:for-each select="//REQUEST/AFTER/APPLICATION[@APPLICATIONNUMBER=$pAppNumber]/CUSTOMERROLE[@CUSTOMERNUMBER=$vCn][@CUSTOMERVERSIONNUMBER=$vCvn]/ALIAS[@ALIASTYPE_TYPE_AL]">
					<xsl:if test="user:HasDataChanged() = 0">
						<xsl:element name="REPROCESS_CREDIT_CHECK">
							<xsl:attribute name="APPLICATION"><xsl:value-of select="$pAppNumber"/></xsl:attribute>
							<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="@CUSTOMERNUMBER"/></xsl:attribute>
							<xsl:element name="ADDED">
								<xsl:attribute name="ID">ALIAS</xsl:attribute>
							</xsl:element>
							<xsl:value-of select="user:SetDataChanged()"/>
						</xsl:element>
					</xsl:if>
				</xsl:for-each>
			</xsl:otherwise>
		</xsl:choose>
		<!-- IK 23/06/2006 EP767 new ALIAS: end -->
		
		<xsl:if test="user:HasDataChanged() = 1">
			<xsl:call-template name="CreateTask">
				<xsl:with-param name="task">TMReprocessCreditCheck</xsl:with-param>
			</xsl:call-template>
		</xsl:if>
	</xsl:template>
	<!--END: Reprocess credit check-->
	<!--START: Rescore credit check-->
	<xsl:template name="ReScoreCreditCheck">
		<xsl:param name="pAppNumber"/>
		<xsl:variable name="vCn">
			<xsl:value-of select="@CUSTOMERNUMBER"/>
		</xsl:variable>
		<xsl:variable name="vCvn">
			<xsl:value-of select="@CUSTOMERVERSIONNUMBER"/>
		</xsl:variable>

		<xsl:variable name="vEmploymentStatus_After">
		<xsl:value-of select="//REQUEST/AFTER/APPLICATION[@APPLICATIONNUMBER=$pAppNumber]/CUSTOMERROLE/CUSTOMERVERSION[@CUSTOMERNUMBER=$vCn][@CUSTOMERVERSIONNUMBER=$vCvn]/EMPLOYMENT/@EMPLOYMENTSTATUS"/>
		</xsl:variable>
		<xsl:variable name="vEmploymentStatus_Before"><xsl:value-of select="EMPLOYMENT[@CUSTOMERVERSIONNUMBER=$vCvn]/@EMPLOYMENTSTATUS"/></xsl:variable>
		<xsl:if test="$vEmploymentStatus_Before != $vEmploymentStatus_After">
			<xsl:if test="user:HasDataChanged() = 0">
				<xsl:element name="RESCORE_CREDIT_CHECK">
					<xsl:attribute name="APPLICATION"><xsl:value-of select="$pAppNumber"/></xsl:attribute>
					<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="@CUSTOMERNUMBER"/></xsl:attribute>
					<xsl:element name="CHANGED">
						<xsl:attribute name="ID">EMPLOYMENTSTATUS</xsl:attribute>
						<xsl:attribute name="before"><xsl:value-of select="$vEmploymentStatus_Before"/></xsl:attribute>
						<xsl:attribute name="after"><xsl:value-of select="$vEmploymentStatus_After"/></xsl:attribute>
					</xsl:element>
					<xsl:value-of select="user:SetDataChanged()"/>
				</xsl:element>
			</xsl:if>
		</xsl:if>
		<xsl:variable name="vDateLeftOrCeasedTrading_After">
			<xsl:value-of select="//REQUEST/AFTER/APPLICATION[@APPLICATIONNUMBER=$pAppNumber]/CUSTOMERROLE/CUSTOMERVERSION[@CUSTOMERNUMBER=$vCn][@CUSTOMERVERSIONNUMBER=$vCvn]/EMPLOYMENT/@DATELEFTORCEASEDTRADING"/>
		</xsl:variable>
		<xsl:variable name="vDateLeftOrCeasedTrading_Before"><xsl:value-of select="EMPLOYMENT[@CUSTOMERVERSIONNUMBER=$vCvn]/@DATELEFTORCEASEDTRADING"/></xsl:variable>
		<xsl:if test="$vDateLeftOrCeasedTrading_Before != $vDateLeftOrCeasedTrading_After">
			<xsl:if test="user:HasDataChanged() = 0">
				<xsl:element name="RESCORE_CREDIT_CHECK">
					<xsl:attribute name="APPLICATION"><xsl:value-of select="$pAppNumber"/></xsl:attribute>
					<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="@CUSTOMERNUMBER"/></xsl:attribute>
					<xsl:element name="CHANGED">
						<xsl:attribute name="ID">DATELEFTORCEASEDTRADING</xsl:attribute>
						<xsl:attribute name="before"><xsl:value-of select="$vDateLeftOrCeasedTrading_Before"/></xsl:attribute>
						<xsl:attribute name="after"><xsl:value-of select="$vDateLeftOrCeasedTrading_After"/></xsl:attribute>
					</xsl:element>
					<xsl:value-of select="user:SetDataChanged()"/>
				</xsl:element>
			</xsl:if>
		</xsl:if>
		<xsl:variable name="vDateStartedOrEstablished_After">
			<xsl:value-of select="//REQUEST/AFTER/APPLICATION[@APPLICATIONNUMBER=$pAppNumber]/CUSTOMERROLE/CUSTOMERVERSION[@CUSTOMERNUMBER=$vCn][@CUSTOMERVERSIONNUMBER=$vCvn]/EMPLOYMENT/@DATESTARTEDORESTABLISHED"/>
		</xsl:variable>
		<xsl:variable name="vDateStartedOrEstablished_Before"><xsl:value-of select="EMPLOYMENT[@CUSTOMERVERSIONNUMBER=$vCvn]/@DATESTARTEDORESTABLISHED"/></xsl:variable>
		<xsl:if test="$vDateStartedOrEstablished_Before != $vDateStartedOrEstablished_After">
			<xsl:if test="user:HasDataChanged() = 0">
				<xsl:element name="RESCORE_CREDIT_CHECK">
					<xsl:attribute name="APPLICATION"><xsl:value-of select="$pAppNumber"/></xsl:attribute>
					<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="@CUSTOMERNUMBER"/></xsl:attribute>
					<xsl:element name="CHANGED">
						<xsl:attribute name="ID">DATESTARTEDORESTABLISHED</xsl:attribute>
						<xsl:attribute name="before"><xsl:value-of select="$vDateStartedOrEstablished_Before"/></xsl:attribute>
						<xsl:attribute name="after"><xsl:value-of select="$vDateStartedOrEstablished_After"/></xsl:attribute>
					</xsl:element>
					<xsl:value-of select="user:SetDataChanged()"/>
				</xsl:element>
			</xsl:if>
		</xsl:if>
				
		<xsl:variable name="vAllowableAnnualIncome_After">
			<xsl:value-of select="//REQUEST/AFTER/APPLICATION[@APPLICATIONNUMBER=$pAppNumber]/CUSTOMERROLE[@CUSTOMERNUMBER=$vCn][@CUSTOMERVERSIONNUMBER=$vCvn]/INCOMESUMMARY[@CUSTOMERVERSIONNUMBER=$vCvn]/@ALLOWABLEANNUALINCOME"/>
		</xsl:variable>
		<xsl:variable name="vAllowableAnnualIncome_Before"><xsl:value-of select="../INCOMESUMMARY[@CUSTOMERVERSIONNUMBER=$vCvn]/@ALLOWABLEANNUALINCOME"/></xsl:variable>
		<xsl:if test="$vAllowableAnnualIncome_Before != $vAllowableAnnualIncome_After">
			<xsl:if test="user:HasDataChanged() = 0">
				<xsl:element name="RESCORE_CREDIT_CHECK">
					<xsl:attribute name="APPLICATION"><xsl:value-of select="$pAppNumber"/></xsl:attribute>
					<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="@CUSTOMERNUMBER"/></xsl:attribute>
					<xsl:element name="CHANGED">
						<xsl:attribute name="ID">ALLOWABLEANNUALINCOME</xsl:attribute>
						<xsl:attribute name="before"><xsl:value-of select="$vAllowableAnnualIncome_Before"/></xsl:attribute>
						<xsl:attribute name="after"><xsl:value-of select="$vAllowableAnnualIncome_After"/></xsl:attribute>
					</xsl:element>
					<xsl:value-of select="user:SetDataChanged()"/>
				</xsl:element>
			</xsl:if>
		</xsl:if>
				
		<xsl:variable name="vNatureOfLoan_After">
			<xsl:value-of select="//REQUEST/AFTER/APPLICATION[@APPLICATIONNUMBER=$pAppNumber]/@NATUREOFLOAN"/>
		</xsl:variable>
		<xsl:variable name="vNatureOfLoan_Before"><xsl:value-of select="../../../APPLICATION/@NATUREOFLOAN"/></xsl:variable>
		<xsl:if test="$vNatureOfLoan_Before != $vNatureOfLoan_After">
			<xsl:if test="user:HasDataChanged() = 0">
				<xsl:element name="RESCORE_CREDIT_CHECK">
					<xsl:attribute name="APPLICATION"><xsl:value-of select="$pAppNumber"/></xsl:attribute>
					<xsl:element name="CHANGED">
						<xsl:attribute name="ID">NATUREOFLOAN</xsl:attribute>
						<xsl:attribute name="before"><xsl:value-of select="$vNatureOfLoan_Before"/></xsl:attribute>
						<xsl:attribute name="after"><xsl:value-of select="$vNatureOfLoan_After"/></xsl:attribute>
					</xsl:element>
					<xsl:value-of select="user:SetDataChanged()"/>
				</xsl:element>
			</xsl:if>
		</xsl:if>
		<!--Use the ActiveQuoteNumber if there is no AcceptedQuoteNumber	-->
		<xsl:variable name="vAcceptedQuotationNumber_Before">
			<xsl:value-of select="../../../APPLICATION/@ACCEPTEDQUOTENUMBER"/>
		</xsl:variable>
		<xsl:variable name="vActiveQuotationNumber_Before">
			<xsl:value-of select="../../../APPLICATION/@ACTIVEQUOTENUMBER"/>
		</xsl:variable>

		<xsl:variable name="vMortgageSubQuoteNumber_Before">
			<xsl:choose>
				<xsl:when test="string-length($vAcceptedQuotationNumber_Before) != 0">
					<xsl:value-of select="../../QUOTATION[@QUOTATIONNUMBER=$vAcceptedQuotationNumber_Before]/@MORTGAGESUBQUOTENUMBER"/>
				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of select="../../QUOTATION[@QUOTATIONNUMBER=$vActiveQuotationNumber_Before]/@MORTGAGESUBQUOTENUMBER"/>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:variable>

		<xsl:variable name="vQuotationNumber_Before">
			<xsl:choose>
				<xsl:when test="string-length(vAcceptedQuotationNumber_Before) != 0">
					<xsl:value-of select="$vAcceptedQuotationNumber_Before"/>
				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of select="$vActiveQuotationNumber_Before"/>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:variable>
		
		<!-- The AcceptedQuote numbers in the before and after XML will not always match -->
		<xsl:variable name="vAcceptedQuotationNumber_After">
			<xsl:value-of select="//REQUEST/AFTER/APPLICATION/@ACCEPTEDQUOTENUMBER"/>
		</xsl:variable>
		<xsl:variable name="vActiveQuotationNumber_After">
			<xsl:value-of select="//REQUEST/AFTER/APPLICATION/@ACTIVEQUOTENUMBER"/>
		</xsl:variable>
		<xsl:variable name="vMortgageSubQuoteNumber_After">
			<xsl:choose>
				<xsl:when test="string-length(vAcceptedQuotationNumber_After) != 0">
					<xsl:value-of select="//REQUEST/AFTER/APPLICATION/QUOTATION[@QUOTATIONNUMBER=$vAcceptedQuotationNumber_After]/@MORTGAGESUBQUOTENUMBER"/>
				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of select="//REQUEST/AFTER/APPLICATION/QUOTATION[@QUOTATIONNUMBER=$vActiveQuotationNumber_After]/@MORTGAGESUBQUOTENUMBER"/>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:variable>
				
		<xsl:variable name="vQuotationNumber_After">
			<xsl:choose>
				<xsl:when test="string-length(vAcceptedQuotationNumber_After) != 0">
					<xsl:value-of select="$vAcceptedQuotationNumber_After"/>
				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of select="$vActiveQuotationNumber_After"/>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:variable>
		

		
		<xsl:variable name="vAmountRequested_Before">
			<xsl:value-of select="../../QUOTATION[@QUOTATIONNUMBER=$vQuotationNumber_Before]/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=$vMortgageSubQuoteNumber_Before]/@AMOUNTREQUESTED"/>
		</xsl:variable>
		<xsl:variable name="vAmountRequested_After">
			<xsl:value-of select="//REQUEST/AFTER/APPLICATION/QUOTATION[@QUOTATIONNUMBER=$vQuotationNumber_After]/MORTGAGESUBQUOTE[@MORTGAGESUBQUOTENUMBER=$vMortgageSubQuoteNumber_After]/@AMOUNTREQUESTED"/>
		</xsl:variable>
		
		<xsl:if test="$vAmountRequested_Before != $vAmountRequested_After">
			<xsl:if test="user:HasDataChanged() = 0">
				<xsl:element name="RESCORE_CREDIT_CHECK">
					<xsl:attribute name="APPLICATION"><xsl:value-of select="$pAppNumber"/></xsl:attribute>
					<xsl:attribute name="QUOTATIONNUMBER"><xsl:value-of select="$vQuotationNumber_Before"/></xsl:attribute>
					<xsl:attribute name="MORTGAGESUBQUOTE"><xsl:value-of select="$vMortgageSubQuoteNumber_Before"/></xsl:attribute>
	
					<xsl:element name="CHANGED">
						<xsl:attribute name="ID">AMOUNTREQUESTED</xsl:attribute>
						<xsl:attribute name="before"><xsl:value-of select="$vAmountRequested_Before"/></xsl:attribute>
						<xsl:attribute name="after"><xsl:value-of select="$vAmountRequested_After"/></xsl:attribute>
					</xsl:element>
					<xsl:value-of select="user:SetDataChanged()"/>
				</xsl:element>
			</xsl:if>
		</xsl:if>
		
		<xsl:variable name="vSelfCertInd_After">
			<xsl:value-of select="//REQUEST/AFTER/APPLICATION[@APPLICATIONNUMBER=$pAppNumber]/@SELFCERTIND"/>
		</xsl:variable>
		<xsl:variable name="vSelfCertInd_Before"><xsl:value-of select="../../../APPLICATION/@SELFCERTIND"/></xsl:variable>
		<xsl:if test="$vSelfCertInd_Before != $vSelfCertInd_After">
			<xsl:if test="user:HasDataChanged() = 0">
				<xsl:element name="RESCORE_CREDIT_CHECK">
					<xsl:attribute name="APPLICATION"><xsl:value-of select="$pAppNumber"/></xsl:attribute>
					<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="@CUSTOMERNUMBER"/></xsl:attribute>
					<xsl:element name="CHANGED">
						<xsl:attribute name="ID">SELFCERTIND</xsl:attribute>
						<xsl:attribute name="before"><xsl:value-of select="$vSelfCertInd_Before"/></xsl:attribute>
						<xsl:attribute name="after"><xsl:value-of select="$vSelfCertInd_After"/></xsl:attribute>
					</xsl:element>
					<xsl:value-of select="user:SetDataChanged()"/>
				</xsl:element>
			</xsl:if>
		</xsl:if>
		
		<xsl:if test="user:HasDataChanged() = 1">
			<xsl:call-template name="CreateTask">
				<xsl:with-param name="task">TMRescoreCreditCheck</xsl:with-param>
			</xsl:call-template>
			<!--EP615: removed SetDataChanged()-->
		</xsl:if>
	</xsl:template>
	<!--END: Rescore credit check-->
	<!--START: CreateTask-->
	<xsl:template name="CreateTask">
		<xsl:param name="task"/>
		<!-- Return case task ID -->
		<xsl:element name="CASETASK">
			<xsl:attribute name="TASKID"><xsl:value-of select="//REQUEST/BEFORE/APPLICATION/GLOBALPARAMETER[@NAME=$task]/@STRING"/></xsl:attribute>
		</xsl:element>
		<!--<xsl:message terminate="yes"/>-->
	</xsl:template>
	<!--END: CreateTask-->
	<!-- SR 11/03/2007: EP2_1828 -->
	<xsl:template name="IntroducerFeeChanged">
		<xsl:param name="pFeeType"/>
		<xsl:param name="pFeeSeqNo"/>
		<xsl:param name="pFeeAmount"/>
		<xsl:param name="pApplicationNumber"/>
		<xsl:choose>
			<xsl:when test="/REQUEST/BEFORE/APPLICATION[@APPLICATIONNUMBER=$pApplicationNumber]/INTRODUCERFEE[@FEESEQUENCENUMBER=$pFeeSeqNo]">
				<xsl:for-each select="/REQUEST/BEFORE/APPLICATION[@APPLICATIONNUMBER=$pApplicationNumber]/INTRODUCERFEE[@FEESEQUENCENUMBER=$pFeeSeqNo]">
					<xsl:variable name="vFeeType">
						<xsl:value-of select="@FEETYPE"/>
					</xsl:variable>
					<xsl:variable name="vFeeAmount">
						<xsl:value-of select="@FEEAMOUNT"/>
					</xsl:variable>
					<xsl:if test="$pFeeType != $vFeeType or $pFeeAmount != $vFeeAmount">
						<xsl:value-of select="user:SetDataChanged()"/>						
					</xsl:if>				
				</xsl:for-each>
			</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select="user:SetDataChanged()"/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<!-- SR 11/03/2007: EP2_1828 - End -->	
	
</xsl:stylesheet>
