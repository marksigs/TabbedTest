<?xml version="1.0" encoding="UTF-8"?>
<!--
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Author		Date			Description
PE			10/11/2006	EP2_24 - Created / Xit2 Web Service
PE			11/12/2006	EP2_414 - AssignFirmID should be blank
PE			16/12/2006	EP2_533 - Problems with the data sent to Xit2 in the Valuation Instruction
PE			06/04/2007	EP2_2272 - Format telephone numbers. (trim to 25 chars to match xsd)
====================================================================================
--> 
<!--=================================================================================================-->
<xsl:stylesheet version="2.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:fn="http://www.w3.org/2005/02/xpath-functions" xmlns:xdt="http://www.w3.org/2005/02/xpath-datatypes">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<!--=================================================================================================-->
	<xsl:template match="/">
		<xsl:apply-templates select="instruction"/>
	</xsl:template>
	<!--=================================================================================================-->
	<xsl:template match="instruction">
		<Instruction xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
			<xsl:attribute name="xsi:noNamespaceSchemaLocation"><xsl:text>C:\john\vex interface\xml\InsertInstruction.xsd</xsl:text></xsl:attribute>
			<InstructionRef>
				<xsl:value-of select="@applicationnumber"/>
				<xsl:value-of select="@instructionref"/>
			</InstructionRef>
			<InstructionRef1/>
			<InstructionRef2/>
			<InstructionRef3/>
			<InstructionDX/>
			<SourceGUID>
				<xsl:value-of select="@sourceguid"/>
			</SourceGUID>
			<WhenCreated>1900-01-01T00:00:00</WhenCreated>
			<ReadyToSend>1</ReadyToSend>
			<IsSent>0</IsSent>
			<WhenSent>1900-01-01T00:00:00</WhenSent>
			<xsl:apply-templates select="insertinstruction"/>
		</Instruction>
	</xsl:template>
	<!--=================================================================================================-->
	<xsl:template match="insertinstruction">
		<InsertInstruction>
			<InstructingSource>LENDERFM</InstructingSource>
			<InstructionDate>
				<xsl:value-of select="@instructiondate"/>
			</InstructionDate>
			<StatusCode>
				<xsl:value-of select="@statuscode"/>
			</StatusCode>
			<ValuationReqd>
				<xsl:if test="@natureofloan=1">
					<xsl:value-of select="@notbuytolet"/>
				</xsl:if>
				<xsl:if test="@natureofloan!=1">
					<xsl:value-of select="@buytolet"/>
				</xsl:if>
			</ValuationReqd>
			<MortgageType>
				<xsl:value-of select="@mortgagetype"/>
			</MortgageType>
			<MortgageReqd>
				<xsl:value-of select="@mortgagereqd"/>
			</MortgageReqd>
			<PurchasePrice>
				<xsl:value-of select="@purchaseprice"/>
			</PurchasePrice>
			<ApproxValue>
				<xsl:value-of select="@approxvalue"/>
			</ApproxValue>
			<FeePaid>
				<xsl:value-of select="@feepaid"/>
			</FeePaid>
			<Currency>0</Currency>
			<GroundRent>
				<xsl:value-of select="@groundrent"/>
			</GroundRent>
			<LeaseHoldTerm>
				<xsl:value-of select="@leaseholdterm"/>
			</LeaseHoldTerm>
			<Tenure>
				<xsl:value-of select="@tenure"/>
			</Tenure>
			<xsl:call-template name="packager"/>
			<xsl:call-template name="pmanager"/>
			<xsl:call-template name="lender"/>
			<xsl:call-template name="return"/>
			<xsl:call-template name="assign"/>
			<xsl:call-template name="new"/>
			<xsl:call-template name="borrowers"/>
			<xsl:call-template name="ea"/>
			<xsl:call-template name="access"/>
			<CommentInstruction/>
		</InsertInstruction>
	</xsl:template>
	<!--=================================================================================================-->
	<xsl:template name="packager">
		<Packager>
			<PackagerFirm>
				<PackagerFirmName/>
				<PackagerFirmID/>			
				<PackagerFirmContact>
					<Phone>
						<Number/>
						<Type>Other</Type>
					</Phone>
					<Fax>
						<Number/>
						<Type>Other</Type>
					</Fax>
					<Email/>
				</PackagerFirmContact>
			</PackagerFirm>
			<PackagerProcessor>
				<PackagerProcessorID/>
				<PackagerProcessorName>
					<FullName/>
					<Title/>
					<ForeName/>
					<MiddleNames/>
					<LastName/>
				</PackagerProcessorName>
				<PackagerProcessorContact>
					<Phone>
						<Number/>
						<Type>Other</Type>
					</Phone>
					<Fax>
						<Number/>
						<Type>Other</Type>
					</Fax>
					<Email/>
				</PackagerProcessorContact>
			</PackagerProcessor>
		</Packager>
	</xsl:template>
	<!--=================================================================================================-->
	<xsl:template name="pmanager">
		<PManager>
			<PManagerFirm>
				<PManagerFirmName/>
				<PManagerFirmID/>
				<PManagerFirmContact>
					<Phone>
						<Number/>
						<Type>Other</Type>
					</Phone>
					<Fax>
						<Number/>
						<Type>Other</Type>
					</Fax>
					<Email/>
				</PManagerFirmContact>
			</PManagerFirm>
			<PManagerProcessor>
				<PManagerProcessorID/>
				<PManagerProcessorName>
					<FullName/>
					<Title/>
					<ForeName/>
					<MiddleNames/>
					<LastName/>
				</PManagerProcessorName>
				<PManagerProcessorContact>
					<Phone>
						<Number/>
						<Type>Other</Type>
					</Phone>
					<Fax>
						<Number/>
						<Type>Other</Type>
					</Fax>
					<Email/>
				</PManagerProcessorContact>
			</PManagerProcessor>
		</PManager>
	</xsl:template>
	<!--=================================================================================================-->
	<xsl:template name="lender">
		<Lender>
			<LenderFirm>
				<LenderFirmName>dbm</LenderFirmName>
				<LenderFirmID>112961</LenderFirmID>
				<LenderFirmContact>
					<Phone>
						<Number/>
						<Type>Other</Type>
					</Phone>
					<Fax>
						<Number/>
						<Type>Other</Type>
					</Fax>
					<Email/>
				</LenderFirmContact>
			</LenderFirm>
			<LenderProcessor>
				<LenderProcessorID/>
				<LenderProcessorName>
					<FullName/>
					<Title/>
					<ForeName/>
					<MiddleNames/>
					<LastName/>
				</LenderProcessorName>
				<LenderProcessorContact>
					<Phone>
						<Number/>
						<Type>Other</Type>
					</Phone>
					<Fax>
						<Number/>
						<Type>Other</Type>
					</Fax>
					<Email/>
				</LenderProcessorContact>
			</LenderProcessor>
		</Lender>
	</xsl:template>
	<!--=================================================================================================-->
	<xsl:template name="return">
		<Return>
			<ReturnFirm>
				<ReturnFirmName>dbm</ReturnFirmName>
				<ReturnFirmID>112961</ReturnFirmID>
			</ReturnFirm>
			<!--
			<ReturnProcessor>
				<ReturnProcessorName>
					<FullName/>
					<Title/>
					<ForeName/>
					<MiddleNames/>
					<LastName/>
				</ReturnProcessorName>
				<ReturnProcessorContact>
					<Phone>
						<Number/>
						<Type>Other</Type>
					</Phone>
					<Fax>
						<Number/>
						<Type>Other</Type>
					</Fax>
					<Email/>
				</ReturnProcessorContact>
			</ReturnProcessor>
			-->
		</Return>
	</xsl:template>
	<!--=================================================================================================-->
	<xsl:template name="assign">
		<Assign>
			<AssignFirm>
				<AssignFirmName/>
				<AssignFirmID></AssignFirmID>
				<AssignFirmAddress>
					<HouseName/>
					<HouseNumber/>
					<Street/>
					<District/>
					<Area/>
					<City/>
					<County/>
					<Country/>
					<PostCode/>
				</AssignFirmAddress>
			</AssignFirm>
			<!--
			<AssignProcessor>
				<AssignProcessorID/>
				<AssignProcessorName>
					<FullName/>
					<Title/>
					<ForeName/>
					<MiddleNames/>
					<LastName/>
				</AssignProcessorName>
				<AssignProcessorContact>
					<Phone>
						<Number/>
						<Type>Other</Type>
					</Phone>
					<Fax>
						<Number/>
						<Type>Other</Type>
					</Fax>
					<Email/>
				</AssignProcessorContact>
			</AssignProcessor>
			-->
		</Assign>
	</xsl:template>
	<!--=================================================================================================-->
	<xsl:template name="new">
		<New>
			<NewAddress>
				<HouseName>
					<xsl:value-of select="/instruction/insertinstruction/new/newaddress/@housename"/>
				</HouseName>
				<HouseNumber>
					<xsl:value-of select="/instruction/insertinstruction/new/newaddress/@housenumber"/>
				</HouseNumber>
				<Street>
					<xsl:value-of select="/instruction/insertinstruction/new/newaddress/@street"/>
				</Street>
				<District>
					<xsl:value-of select="/instruction/insertinstruction/new/newaddress/@district"/>
				</District>
				<Area/>
				<City>
					<xsl:value-of select="/instruction/insertinstruction/new/newaddress/@city"/>
				</City>
				<County>
					<xsl:value-of select="/instruction/insertinstruction/new/newaddress/@county"/>
				</County>
				<Country>
					<xsl:value-of select="/instruction/insertinstruction/new/newaddress/@country"/>
				</Country>
				<PostCode>
					<xsl:value-of select="/instruction/insertinstruction/new/newaddress/@postcode"/>
				</PostCode>
			</NewAddress>
			<!--
			<NewName>
				<FullName/>
				<Title/>
				<ForeName/>
				<MiddleNames/>
				<LastName/>
			</NewName>
			<NewContact>
				<Phone>
					<Number/>
					<Type>Other</Type>
				</Phone>
				<Fax>
					<Number/>
					<Type>Other</Type>
				</Fax>
				<Email/>
			</NewContact>
			-->
		</New>
	</xsl:template>
	<!--=================================================================================================-->
	<xsl:template name="borrowers">
		<Borrowers>
			<Borrower1Name>
				<FullName>
					<xsl:value-of select="/instruction/insertinstruction/borrowers/borrower1/@fullname"/>
				</FullName>
				<Title>
					<xsl:value-of select="/instruction/insertinstruction/borrowers/borrower1/@title"/>
				</Title>
				<ForeName>
					<xsl:value-of select="/instruction/insertinstruction/borrowers/borrower1/@forename"/>
				</ForeName>
				<MiddleNames>
					<xsl:value-of select="/instruction/insertinstruction/borrowers/borrower1/@middlenames"/>
				</MiddleNames>
				<LastName>
					<xsl:value-of select="/instruction/insertinstruction/borrowers/borrower1/@lastname"/>
				</LastName>
			</Borrower1Name>
			<xsl:if test="instruction/insertinstruction/borrowers/borrower2/@fullname!=''">
				<Borrower2Name>
					<FullName>
						<xsl:value-of select="/instruction/insertinstruction/borrowers/borrower2/@fullname"/>
					</FullName>
					<Title>
						<xsl:value-of select="/instruction/insertinstruction/borrowers/borrower2/@title"/>
					</Title>
					<ForeName>
						<xsl:value-of select="/instruction/insertinstruction/borrowers/borrower2/@forename"/>
					</ForeName>
					<MiddleNames>
						<xsl:value-of select="/instruction/insertinstruction/borrowers/borrower2/@middlenames"/>
					</MiddleNames>
					<LastName>
						<xsl:value-of select="/instruction/insertinstruction/borrowers/borrower2/@lastname"/>
					</LastName>
				</Borrower2Name>
			</xsl:if>
			<BorrowersContact>
				<Phone>
					<Number>
						<xsl:value-of select="substring(/instruction/insertinstruction/borrowers/borrowerscontact[@usage='Home']/@phone,1,25)"/>
					</Number>
					<Type>Other</Type>
				</Phone>
				<Fax>
					<Number>
						<xsl:value-of select="substring(/instruction/insertinstruction/borrowers/borrowerscontact[@usage='Fax']/@phone,1,25)"/>
					</Number>
					<Type>Other</Type>
				</Fax>
				<Email>
					<xsl:value-of select="/instruction/insertinstruction/borrowers/borrowerscontact/@email"/>
				</Email>
			</BorrowersContact>
			<BorrowersAddress>
				<HouseName>
					<xsl:value-of select="/instruction/insertinstruction/borrowers/borrowersaddress/@housename"/>
				</HouseName>
				<HouseNumber>
					<xsl:value-of select="/instruction/insertinstruction/borrowers/borrowersaddress/@housenumber"/>
				</HouseNumber>
				<Street>
					<xsl:value-of select="/instruction/insertinstruction/borrowers/borrowersaddress/@street"/>
				</Street>
				<District>
					<xsl:value-of select="/instruction/insertinstruction/borrowers/borrowersaddress/@district"/>
				</District>
				<Area/>
				<City>
					<xsl:value-of select="/instruction/insertinstruction/borrowers/borrowersaddress/@city"/>
				</City>
				<County>
					<xsl:value-of select="/instruction/insertinstruction/borrowers/borrowersaddress/@county"/>
				</County>
				<Country>
					<xsl:value-of select="/instruction/insertinstruction/borrowers/borrowersaddress/@country"/>
				</Country>
				<PostCode>
					<xsl:value-of select="/instruction/insertinstruction/borrowers/borrowersaddress/@postcode"/>
				</PostCode>
			</BorrowersAddress>
		</Borrowers>
	</xsl:template>
	<!--=================================================================================================-->
	<xsl:template name="ea">
		<EA>
			<EAFirm>
				<EAFirmName>
					<xsl:value-of select="/instruction/insertinstruction/ea/eafirm/@eafirmname"/>
				</EAFirmName>
				<EAFirmContact>
					<Phone>
						<Number>
							<xsl:value-of select="substring(/instruction/insertinstruction/ea/eafirm/@phone,1,25)"/>
						</Number>
						<Type>Other</Type>
					</Phone>
					<Fax>
						<Number>
							<xsl:value-of select="substring(/instruction/insertinstruction/ea/eafirm/@fax,1,25)"/>
						</Number>
						<Type>Other</Type>
					</Fax>
					<Email>
						<xsl:value-of select="/instruction/insertinstruction/ea/eafirm/@email"/>
					</Email>
				</EAFirmContact>
			</EAFirm>
		</EA>
	</xsl:template>
	<!--=================================================================================================-->
	<xsl:template name="access">
		<Access>
			<AccessName>
				<FullName>
					<xsl:value-of select="/instruction/insertinstruction/access/@fullname"/>
				</FullName>
				<Title/>
				<ForeName/>
				<MiddleNames/>
				<LastName/>
			</AccessName>
			<AccessContact>
				<Phone>
					<Number>
						<xsl:value-of select="substring(/instruction/insertinstruction/access/@phone,1,25)"/>
					</Number>
					<Type>Other</Type>
				</Phone>
				<Fax>
					<Number/>
					<Type>Other</Type>
				</Fax>
				<Email/>
			</AccessContact>
		</Access>
	</xsl:template>
	<!--=================================================================================================-->
</xsl:stylesheet>
<!--=================================================================================================-->
