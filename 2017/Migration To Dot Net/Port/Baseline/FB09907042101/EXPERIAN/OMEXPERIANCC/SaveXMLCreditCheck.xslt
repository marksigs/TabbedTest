<?xml version="1.0" encoding="UTF-8"?>
<!--==============================XML Document Control=============================
Description: SaveXMLCreditCheck

History:

Version Author     Date       Description
01.01	LDM			15/02/2006 EP6  Epsom. Updated to deal with new layout for Cut from mars v9
01.02   LDM			11/04/06     EP373 New besopke fields being returned in the bes1 block
01.03   LDM			12/04/06     EP373 New field BX12_UKMOSAIC being returned.
01.04   LDM			19/04/06     EP339 fix bx06_cmlDetected
01.05	LDM			25/04/06     EP456 New fields BES1_SECRENT_MISS_L12M, BES1_SECRENT_MISS_L6M, BES1_SECRENT_WCS
01.06	SAB		09/05/06	   EP414 Updated to find the correct Current Afforability Index field.
01.07	SAB		11/05/06	   EP483 Updated to pass the Nature of Occupancy
01.08   LDM			24/05/06     EP597 Address Targetting issue was saving blank addresses
01.09   AW			09/11/06     EP1246 New BureauData BES1 fields
================================================================================-->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<xsl:variable name="Response" select="/GEODS"/>
	<xsl:variable name="CreditCheckRoot" select="$Response/REQUEST"/>
	<xsl:variable name="OmigaData" select="$Response/REQUEST/OMIGADATA"/>
	<xsl:variable name="PrimaryApplicant" select="$OmigaData/PRIMARYAPPLICANT"/>
	<xsl:variable name="SecondApplicant" select="$OmigaData/SECONDAPPLICANT"/>
	<xsl:variable name="TargetingData" select="$OmigaData/TARGETINGDATA"/>
	<xsl:variable name="AddressMap" select="$TargetingData/ADDRESSMAP"/>
	<xsl:variable name="ExpAddress" select="$CreditCheckRoot/ADDR-UKFMTD"/>
	<xsl:variable name="BDX1" select="$CreditCheckRoot/BDX1"/>
	<xsl:variable name="BX01" select="$CreditCheckRoot/BX01"/>
	<xsl:variable name="BX02" select="$CreditCheckRoot/BX02"/>
	<xsl:variable name="BX03" select="$CreditCheckRoot/BX03"/>
	<xsl:variable name="BX04" select="$CreditCheckRoot/BX04"/>
	<xsl:variable name="BX05" select="$CreditCheckRoot/BX05"/>
	<xsl:variable name="BX06" select="$CreditCheckRoot/BX06"/>
	<xsl:variable name="BX07" select="$CreditCheckRoot/BX07"/>
	<xsl:variable name="BX08" select="$CreditCheckRoot/BX08"/>
	<xsl:variable name="BX09" select="$CreditCheckRoot/BX09"/>	
	<xsl:variable name="BX10" select="$CreditCheckRoot/BX10"/>
	<xsl:variable name="BX11" select="$CreditCheckRoot/BX11"/>
	<xsl:variable name="BX12" select="$CreditCheckRoot/BX12"/>
	<xsl:variable name="BX13" select="$CreditCheckRoot/BX13"/>
	<xsl:variable name="BX14" select="$CreditCheckRoot/BX14"/>
	<xsl:variable name="BX15" select="$CreditCheckRoot/BX15"/>
	<xsl:variable name="BX16" select="$CreditCheckRoot/BX16"/>
	<xsl:variable name="BX17" select="$CreditCheckRoot/BX17"/>
	<xsl:variable name="BX18" select="$CreditCheckRoot/BX18"/>
	<xsl:variable name="DEC1" select="$CreditCheckRoot/DEC1"/>
	<xsl:variable name="DEC7" select="$CreditCheckRoot/DEC7"/>	
	<xsl:variable name="RRP2" select="$CreditCheckRoot/RRP2"/>	
	<xsl:variable name="TPD1" select="$CreditCheckRoot/TPD1_RES"/>		
	<xsl:variable name="BES1" select="$CreditCheckRoot/BES1"/>	
	<xsl:template match="/">
		 <xsl:element name="DATA">
			<xsl:element name="MISC">
				<xsl:element name="APPLICATIONCREDITCHECK">
					<xsl:attribute name="APPLICATIONNUMBER"><xsl:value-of select="$OmigaData/APPLICATION/@APPLICATIONNUMBER"/></xsl:attribute>
					<xsl:attribute name="APPLICATIONFACTFINDNUMBER"><xsl:value-of select="$OmigaData/APPLICATION/@APPLICATIONFACTFINDNUMBER"/></xsl:attribute>
					<xsl:attribute name="USERID"><xsl:value-of select="$Response/@USERID"/></xsl:attribute>
					<xsl:attribute name="SUCCESSINDICATOR">
						<xsl:choose>
							<xsl:when test="$CreditCheckRoot/@success = 'Y'">1</xsl:when>
							<xsl:when test="$CreditCheckRoot/@success = 'N'">0</xsl:when>
						</xsl:choose>
					</xsl:attribute>
					<xsl:attribute name="DATETIME"><xsl:value-of select="$CreditCheckRoot/OMIGADATA/CREDITCHECKDATE/@DATE"/></xsl:attribute>
					<xsl:attribute name="CREDITCHECKREFERENCENUMBER"><xsl:value-of select="$CreditCheckRoot/@EXP_ExperianRef"/></xsl:attribute>
					<xsl:attribute name="TPDOUTCOMECODE"><xsl:value-of select="$TPD1/OUTCOMECODE"/></xsl:attribute>
					<xsl:attribute name="TPDALERTFORASSOCIATESCORE"><xsl:value-of select="$TPD1/OPTOUTCUTOFF"/></xsl:attribute>
					<xsl:attribute name="EXPERIANFUNCTION"><xsl:value-of select="$OmigaData/FUNCTION/@ID"/></xsl:attribute>
				</xsl:element>
				<xsl:element name="CREDITCHECK">
					<xsl:attribute name="ROOTID"><xsl:value-of select="$DEC1/ROOTID"/></xsl:attribute>
					<xsl:attribute name="DECISIONTESTGROUP"><xsl:value-of select="$DEC1/DECTESTGRP"/></xsl:attribute>
					<xsl:attribute name="DECISIONCLASS"><xsl:value-of select="$DEC1/DECCLASS"/></xsl:attribute>
					<xsl:attribute name="DELPHISCORESIGN"><xsl:value-of select="$DEC1/DELPHISCORESIGN"/></xsl:attribute>
					<xsl:attribute name="DELPHISCORE"><xsl:value-of select="$DEC1/DELPHISCORE"/></xsl:attribute>
					<xsl:attribute name="SCORINGVERSION"><xsl:value-of select="$DEC1/SCORING"/></xsl:attribute>
					<xsl:attribute name="POLICYVERSION"><xsl:value-of select="$DEC1/POLICY"/></xsl:attribute>
					<xsl:attribute name="BUSINESSVERSION"><xsl:value-of select="$DEC1/BUSINESS"/></xsl:attribute>
					<xsl:attribute name="APPMANVERSION"><xsl:value-of select="$DEC1/APPMAN"/></xsl:attribute>
					<xsl:attribute name="STRATTEXTID"><xsl:value-of select="$DEC1/STRATEGYTEXTID"/></xsl:attribute>
					<xsl:attribute name="BUREAUFLAG"><xsl:value-of select="$DEC1/BURFLAG"/></xsl:attribute>			
					<xsl:attribute name="QUEUEID"><xsl:value-of select="$DEC1/QUEUEID"/></xsl:attribute>						
					<xsl:attribute name="REASONCODESLENGTH"><xsl:value-of select="$DEC1/RSNCODELEN"/></xsl:attribute>
					<xsl:attribute name="NOSCORENODES"><xsl:value-of select="$DEC1/NUMSCORENODES"/></xsl:attribute>
					<xsl:attribute name="NOLETTERCODES"><xsl:value-of select="$DEC1/NUMLETTERCODES"/></xsl:attribute>
				</xsl:element>
				<xsl:element name="BESPOKEDECISION">
					<xsl:attribute name="RANDOMNUMBER1"><xsl:value-of select="$DEC7/RANNUM1"/></xsl:attribute>
					<xsl:attribute name="RANDOMNUMBER2"><xsl:value-of select="$DEC7/RANNUM2"/></xsl:attribute>
					<xsl:attribute name="RANDOMNUMBER3"><xsl:value-of select="$DEC7/RANNUM3"/></xsl:attribute>
					<xsl:attribute name="RANDOMNUMBER4"><xsl:value-of select="$DEC7/RANNUM4"/></xsl:attribute>
					<xsl:attribute name="RANDOMNUMBER5"><xsl:value-of select="$DEC7/RANNUM5"/></xsl:attribute>
					<xsl:attribute name="RANDOMNUMBER6"><xsl:value-of select="$DEC7/RANNUM6"/></xsl:attribute>
					<xsl:attribute name="RANDOMNUMBER7"><xsl:value-of select="$DEC7/RANNUM7"/></xsl:attribute>
					<xsl:attribute name="RANDOMNUMBER8"><xsl:value-of select="$DEC7/RANNUM8"/></xsl:attribute>
					<xsl:attribute name="RANDOMNUMBER9"><xsl:value-of select="$DEC7/RANNUM9"/></xsl:attribute>
					<xsl:attribute name="RANDOMNUMBER10"><xsl:value-of select="$DEC7/RANNUM10"/></xsl:attribute>
					<xsl:attribute name="ELIGIBILITYINDICATOR"><xsl:value-of select="$DEC7/ELIGIBILITYIND"/></xsl:attribute>
					<xsl:attribute name="DUPLICATECOUNT"><xsl:value-of select="$DEC7/DUPCOUNT"/></xsl:attribute>
					<xsl:attribute name="RISKGRADE"><xsl:value-of select="$DEC7/RISKGRADE"/></xsl:attribute>
					<xsl:attribute name="AFFORDABILITYGRADE"><xsl:value-of select="$DEC7/AFFORDGRADE"/></xsl:attribute>
					<xsl:attribute name="ADVERSEREASONCODE1"><xsl:value-of select="$DEC7/ADREASONC1"/></xsl:attribute>
					<xsl:attribute name="ADVERSEREASONCODE2"><xsl:value-of select="$DEC7/ADREASONC2"/></xsl:attribute>
					<xsl:attribute name="ADVERSEREASONCODE3"><xsl:value-of select="$DEC7/ADREASONC3"/></xsl:attribute>
					<xsl:attribute name="ADVERSEREASONCODE4"><xsl:value-of select="$DEC7/ADREASONC4"/></xsl:attribute>
					<xsl:attribute name="ADVERSEREASONCODE5"><xsl:value-of select="$DEC7/ADREASONC5"/></xsl:attribute>
					<xsl:attribute name="ADVERSEREASONCODE6"><xsl:value-of select="$DEC7/ADREASONC6"/></xsl:attribute>
					<xsl:attribute name="ADVERSEREASONCODE7"><xsl:value-of select="$DEC7/ADREASONC7"/></xsl:attribute>
					<xsl:attribute name="ADVERSEREASONCODE8"><xsl:value-of select="$DEC7/ADREASONC8"/></xsl:attribute>
					<xsl:attribute name="ADVERSEREASONCODE9"><xsl:value-of select="$DEC7/ADREASONC9"/></xsl:attribute>
					<xsl:attribute name="ADVERSEREASONCODE10"><xsl:value-of select="$DEC7/ADREASONC10"/></xsl:attribute>
					<xsl:attribute name="ADVERSEREASONCODE11"><xsl:value-of select="$DEC7/ADREASONC11"/></xsl:attribute>
					<xsl:attribute name="ADVERSEREASONCODE12"><xsl:value-of select="$DEC7/ADREASONC12"/></xsl:attribute>
					<xsl:attribute name="ADVERSEREASONCODE13"><xsl:value-of select="$DEC7/ADREASONC13"/></xsl:attribute>
					<xsl:attribute name="ADVERSEREASONCODE14"><xsl:value-of select="$DEC7/ADREASONC14"/></xsl:attribute>
					<xsl:attribute name="ADVERSEREASONCODE15"><xsl:value-of select="$DEC7/ADREASONC15"/></xsl:attribute>
					<xsl:attribute name="ADVERSEREASONCODE16"><xsl:value-of select="$DEC7/ADREASONC16"/></xsl:attribute>
					<xsl:attribute name="ADVERSEREASONCODE17"><xsl:value-of select="$DEC7/ADREASONC17"/></xsl:attribute>
					<xsl:attribute name="ADVERSEREASONCODE18"><xsl:value-of select="$DEC7/ADREASONC18"/></xsl:attribute>
					<xsl:attribute name="ADVERSEREASONCODE19"><xsl:value-of select="$DEC7/ADREASONC19"/></xsl:attribute>
					<xsl:attribute name="ADVERSEREASONCODE20"><xsl:value-of select="$DEC7/ADREASONC20"/></xsl:attribute>
					<xsl:attribute name="SPAREFIELD1"><xsl:value-of select="$DEC7/SPARE1"/></xsl:attribute>
					<xsl:attribute name="SPAREFIELD2"><xsl:value-of select="$DEC7/SPARE2"/></xsl:attribute>
					<xsl:attribute name="SPAREFIELD3"><xsl:value-of select="$DEC7/SPARE3"/></xsl:attribute>
					<xsl:attribute name="SPAREFIELD4"><xsl:value-of select="$DEC7/SPARE4"/></xsl:attribute>
					<xsl:attribute name="SPAREFIELD5"><xsl:value-of select="$DEC7/SPARE5"/></xsl:attribute>
					<xsl:attribute name="SPAREFIELD6"><xsl:value-of select="$DEC7/SPARE6"/></xsl:attribute>
					<xsl:attribute name="SPAREFIELD7"><xsl:value-of select="$DEC7/SPARE7"/></xsl:attribute>
					<xsl:attribute name="SPAREFIELD8"><xsl:value-of select="$DEC7/SPARE8"/></xsl:attribute>
					<xsl:attribute name="SPAREFIELD9"><xsl:value-of select="$DEC7/SPARE9"/></xsl:attribute>
					<xsl:attribute name="SPAREFIELD10"><xsl:value-of select="$DEC7/SPARE10"/></xsl:attribute>
				</xsl:element>
				<xsl:element name="BXAPPLICATIONBUREAU">
					<xsl:attribute name="BX01_VR_PAFNOTRACE_MACA"><xsl:value-of select="$BX01/NDERLMACA"/></xsl:attribute>
					<xsl:attribute name="BX01_VR_PAFNOTRACE_MAPA"><xsl:value-of select="$BX01/NDERLMAPA"/></xsl:attribute>
					<xsl:attribute name="BX01_VR_PAFNOTRACE_JACA"><xsl:value-of select="$BX01/NDERLJACA"/></xsl:attribute>
					<xsl:attribute name="BX01_VR_PAFNOTRACE_JAPA"><xsl:value-of select="$BX01/NDERLJAPA"/></xsl:attribute>
					<xsl:attribute name="BX01_CURRENTADDRESSNAMECOUNT"><xsl:value-of select="$BX01/EA5U01"/></xsl:attribute>
					<xsl:attribute name="BX01_PREVIOUSADDRESSNAMECOUNT"><xsl:value-of select="$BX01/EA5U02"/></xsl:attribute>
					<xsl:attribute name="BX02_SATISFIEDCCJDETECTED"><xsl:value-of select="$BX02/EA4Q06"/></xsl:attribute>
					<xsl:attribute name="BX03_TOTALUNSECUREDCAISBALANCE"><xsl:value-of select="$BX03/NDECC06"/></xsl:attribute>
					<xsl:attribute name="BX03_TOTALINACTIVECAIS"><xsl:value-of select="$BX03/NDINC03"/></xsl:attribute>
					<xsl:attribute name="BX04_TOTALNDSEARCHES12M"><xsl:value-of select="$BX04/NDPSD11"/></xsl:attribute>
					<xsl:attribute name="BX05_CIFASDETECTED"><xsl:value-of select="$BX05/EA4P01"/></xsl:attribute>
					<xsl:attribute name="BX08_CAPSECAPSNOC"><xsl:value-of select="$BX08/EA4Q07"/></xsl:attribute>
					<xsl:attribute name="BX08_ADDRESSLINKSNOC"><xsl:value-of select="$BX08/EA4Q08"/></xsl:attribute>
					<xsl:attribute name="BX08_ASSOCIATIONSNOC"><xsl:value-of select="$BX08/EA4Q09"/></xsl:attribute>
					<xsl:attribute name="BX08_GAINNOC"><xsl:value-of select="$BX08/EA4Q10"/></xsl:attribute>
					<xsl:attribute name="BX08_DIRECTORSNOC"><xsl:value-of select="$BX08/EA4Q11"/></xsl:attribute>
					<xsl:attribute name="BX08_VOTERSROLLNOC"><xsl:value-of select="$BX08/EA4Q01"/></xsl:attribute>
					<xsl:attribute name="BX08_CCJNOC"><xsl:value-of select="$BX08/EA4Q02"/></xsl:attribute>
					<xsl:attribute name="BX08_CIFASNOC"><xsl:value-of select="$BX08/EA4Q03"/></xsl:attribute>
					<xsl:attribute name="BX08_CAISNOC"><xsl:value-of select="$BX08/EA4Q04"/></xsl:attribute>
					<xsl:attribute name="BX08_OTHERNOC"><xsl:value-of select="$BX08/EA4Q05"/></xsl:attribute>
					<xsl:attribute name="BX09_OPTOUTFLAG"><xsl:value-of select="$BX09/NDOPTOUT"/></xsl:attribute>
					<xsl:attribute name="BX10_PAREFERENCE"><xsl:value-of select="$BX10/NDPA"/></xsl:attribute>
					<xsl:attribute name="BX10_HOUSEHOLDOVERRIDE"><xsl:value-of select="$BX10/NDHHO"/></xsl:attribute>
					<xsl:attribute name="BX10_OPTOUTVALIDFLAG"><xsl:value-of select="$BX10/NDOPTOUTVALID"/></xsl:attribute>
					<xsl:attribute name="BX12_MOSIAC93CODE"><xsl:value-of select="$BX12/EA4T01"/></xsl:attribute>
					<xsl:attribute name="BX12_N_I_MOSAICCODE"><xsl:value-of select="$BX12/EA5T01"/></xsl:attribute>
					<xsl:attribute name="BX12_FINANCIALMOSAICCODE"><xsl:value-of select="$BX12/EA5T02"/></xsl:attribute>
					<xsl:attribute name="BX12_GEODELPHIINDEX"><xsl:value-of select="$BX12/NDG01"/></xsl:attribute>
					<xsl:attribute name="BX12_WEALTHSCORE"><xsl:value-of select="$BX12/EA4N01"/></xsl:attribute>
					<xsl:attribute name="BX12_CHILDSCORE"><xsl:value-of select="$BX12/EA4N02"/></xsl:attribute>
					<xsl:attribute name="BX12_STABILITYSCORE"><xsl:value-of select="$BX12/EA4N03"/></xsl:attribute>
					<xsl:attribute name="BX12_RURABILITYSCORE"><xsl:value-of select="$BX12/EA4N04"/></xsl:attribute>
					<xsl:attribute name="BX12_EDUCATIONSCORE"><xsl:value-of select="$BX12/EA4N05"/></xsl:attribute>
					<xsl:attribute name="BX12_GEODELPHISCORE"><xsl:value-of select="$BX12/NDG02"/></xsl:attribute>
					<xsl:attribute name="BX12_PCAVENUMBEROFCCJSPERHH"><xsl:value-of select="$BX12/NDG03"/></xsl:attribute>
					<xsl:attribute name="BX12_PCPERCENTHHSWITHCCJSLAST36M"><xsl:value-of select="$BX12/NDG04"/></xsl:attribute>
					<xsl:attribute name="BX12_PCAVENOOFCAIS89SPERHH"><xsl:value-of select="$BX12/NDG05"/></xsl:attribute>
					<xsl:attribute name="BX12_PCPERHHSWITH3CAIS8_9S"><xsl:value-of select="$BX12/NDG06"/></xsl:attribute>
					<xsl:attribute name="BX12_PCPERHHSWITHCAIS8_9_L24M"><xsl:value-of select="$BX12/NDG07"/></xsl:attribute>
					<xsl:attribute name="BX12_PCPERHHSWITHDELINQCAIS"><xsl:value-of select="$BX12/NDG08"/></xsl:attribute>
					<xsl:attribute name="BX12_PCPERHHSWORSTCURRSTATUS3"><xsl:value-of select="$BX12/NDG09"/></xsl:attribute>
					<xsl:attribute name="BX12_PCPERHHWRSTSTATUSLAST6M4"><xsl:value-of select="$BX12/NDG10"/></xsl:attribute>
					<xsl:attribute name="BX12_PCPERHHWRSTCURSTATUSREV3"><xsl:value-of select="$BX12/NDG11"/></xsl:attribute>
					<xsl:attribute name="BX12_PCPERHHSWITHLIMIT10K"><xsl:value-of select="$BX12/NDG12"/></xsl:attribute>
					<xsl:attribute name="BX13_SCOREID_2_1"><xsl:value-of select="$BX13/NDSI2-1"/></xsl:attribute>
					<xsl:attribute name="BX13_SCOREID_2_2"><xsl:value-of select="$BX13/NDSI2-2"/></xsl:attribute>
					<xsl:attribute name="BX13_SCOREID_2_3"><xsl:value-of select="$BX13/NDSI2-3"/></xsl:attribute>
					<xsl:attribute name="BX13_POPULATIONINDICATOR"><xsl:value-of select="$BX13/E5S01"/></xsl:attribute>
					<xsl:attribute name="BX13_SUBPOPULATIONINDICATOR"><xsl:value-of select="$BX13/E5S02"/></xsl:attribute>
					<xsl:attribute name="BX13_SCORECARDID1"><xsl:value-of select="$BX13/E5S04-1"/></xsl:attribute>
					<xsl:attribute name="BX13_OPTINSCORE"><xsl:value-of select="$BX13/E5S05-1"/></xsl:attribute>
					<xsl:attribute name="BX13_SCORECARDID2"><xsl:value-of select="$BX13/E5S04-2"/></xsl:attribute>
					<xsl:attribute name="BX13_OPTOUTSCORE"><xsl:value-of select="$BX13/E5S05-2"/></xsl:attribute>
					<xsl:attribute name="BX13_SCORECARDID3"><xsl:value-of select="$BX13/E5S04-3"/></xsl:attribute>
					<xsl:attribute name="BX13_ALERTFORASSOCSCORE"><xsl:value-of select="$BX13/E5S05-3"/></xsl:attribute>
					<xsl:attribute name="BX13_HOUSEHOLDOVERRIDESCORE"><xsl:value-of select="$BX13/NDHHOSCORE"/></xsl:attribute>
					<xsl:attribute name="BX13_IDVALIDATIONSCORENOTAVAIL"><xsl:value-of select="$BX13/NDVALSCORE"/></xsl:attribute>
					<xsl:attribute name="BX14_ADDRESSLINKDATA"><xsl:value-of select="$BX14/NDLNK01"/></xsl:attribute>
					<xsl:attribute name="BX17_VOTERSROLLDOBCOUNT"><xsl:value-of select="$BX17/EA4S01"/></xsl:attribute>
					<xsl:attribute name="BX17_VOTERSROLLDOB"><xsl:value-of select="$BX17/EA4S02"/></xsl:attribute>
					<xsl:attribute name="BX17_CAISDOBCOUNT"><xsl:value-of select="$BX17/EA4S03"/></xsl:attribute>
					<xsl:attribute name="BX17_CAISDOB"><xsl:value-of select="$BX17/EA4S04"/></xsl:attribute>
					<xsl:attribute name="BX17_SEARCHDOBCOUNT"><xsl:value-of select="$BX17/EA4S05"/></xsl:attribute>
					<xsl:attribute name="BX17_SEARCHDOB"><xsl:value-of select="$BX17/EA4S06"/></xsl:attribute>
					<xsl:attribute name="BX17_ALLTYPESDOBCOUNT"><xsl:value-of select="$BX17/EA4S07"/></xsl:attribute>
					<xsl:attribute name="BX17_ALLTYPESDOB"><xsl:value-of select="$BX17/EA4S08"/></xsl:attribute>
					<xsl:attribute name="BX18_IMPAIREDCREDITHISTORY"><xsl:value-of select="$BX18/ND_ICH"/></xsl:attribute>
					<xsl:attribute name="BX18_SECUREDARREARS"><xsl:value-of select="$BX18/ND_SEC_ARR"/></xsl:attribute>
					<xsl:attribute name="BX18_UNSECUREDARREARS"><xsl:value-of select="$BX18/ND_UNSEC_ARR"/></xsl:attribute>
					<xsl:attribute name="BX18_CCJICH"><xsl:value-of select="$BX18/ND_CCJ"/></xsl:attribute>
					<xsl:attribute name="BX18_INDIVIDUALVOLARRANGEMENT"><xsl:value-of select="$BX18/ND_IVA"/></xsl:attribute>
					<xsl:attribute name="BX18_BANKRUPTCY"><xsl:value-of select="$BX18/ND_BANKRUPT"/></xsl:attribute>
					<xsl:attribute name="BX12_UKMOSAIC"><xsl:value-of select="$BX12/UKMOSAIC"/></xsl:attribute>
					<xsl:if test="string($RRP2/CURRENT_AFFORDABILITY_INDEX)">
						<xsl:attribute name="RRP2_CURRENTAFFORDABILITYINDEX"><xsl:value-of select="$RRP2/CURRENT_AFFORDABILITY_INDEX"/></xsl:attribute>
					</xsl:if>
				</xsl:element>			
				<xsl:for-each select="$BDX1">
					<xsl:element name="DETECT">
						<xsl:attribute name="APPLICANTADDRESSSEQUENCE"><xsl:value-of select="@seq"/></xsl:attribute>
						<xsl:attribute name="FRAUDINDEX"><xsl:value-of select="FRAUDINDEX"/></xsl:attribute>
						<!-- <xsl:attribute name="AUTHINDEX"><xsl:value-of select="''"/></xsl:attribute> -->
						<xsl:attribute name="DETECTCREDITSCORE"><xsl:value-of select="DET_CRED_SCR"/></xsl:attribute>
						<xsl:attribute name="DETECTAPPSCORE"><xsl:value-of select="DET_APP_SCR"/></xsl:attribute>
						<xsl:attribute name="NUMBEROFRULES"><xsl:value-of select="NUM_RULES"/></xsl:attribute>
					</xsl:element>
				</xsl:for-each>
				<xsl:element name="ADDRESSBLOCK">
					<xsl:for-each select="$ExpAddress[DETAILCODE != '']">
						<xsl:call-template name="WriteAddressRequiringUpdate">
							<xsl:with-param name="Seq" select="./@seq"/>
						</xsl:call-template>	
					</xsl:for-each>				
				</xsl:element>
			</xsl:element>
			<xsl:element name="CREDITCHECKREASONCODEBLOCK">
				<xsl:for-each select="$OmigaData/CREDITCHECKREASONCODE">
					<xsl:copy-of select="."/>
				</xsl:for-each>
			</xsl:element>
			<xsl:element name="CREDITCHECKLETTERCODEBLOCK">
				<xsl:for-each select="$DEC1/LETTERCODE/CODE">
					<xsl:element name="CREDITCHECKLETTERCODE">
						<xsl:attribute name="SEQUENCENUMBER"><xsl:value-of select="position()"/></xsl:attribute>
						<xsl:attribute name="LETTERCODE"><xsl:value-of select="."/></xsl:attribute>
					</xsl:element>
				</xsl:for-each>
			</xsl:element>
			<xsl:element name="CREDITCHECKSCOREBLOCK">
				<xsl:for-each select="$DEC1/SCORENODE">
					<xsl:element name="CREDITCHECKSCORE">
						<xsl:attribute name="SEQUENCENUMBER"><xsl:value-of select="position()"/></xsl:attribute>
						<xsl:attribute name="SCORENODEID"><xsl:value-of select="./ID"/></xsl:attribute>
						<xsl:attribute name="SCORETESTGROUP"><xsl:value-of select="./TESTGRP"/></xsl:attribute>
						<xsl:attribute name="SCORESIGN"><xsl:value-of select="./SIGN"/></xsl:attribute>
						<xsl:attribute name="SCORE"><xsl:value-of select="./SCORE"/></xsl:attribute>
						<xsl:attribute name="SCORECARDID"><xsl:value-of select="./SCORECARDID"/></xsl:attribute>
					</xsl:element>
				</xsl:for-each>
			</xsl:element>
			<xsl:element name="DETECTRULES">
				<xsl:for-each select="$BDX1">
					<xsl:for-each select="$BDX1//RULES/RULE">
						<xsl:element name="DETECTRULECODE">
							<xsl:attribute name="APPLICANTADDRESSSEQUENCE"><xsl:value-of select="$BDX1/@seq"/></xsl:attribute>
							<xsl:attribute name="SEQUENCENUMBER"><xsl:value-of select="position()"/></xsl:attribute>
							<xsl:attribute name="RULECODE"><xsl:value-of select="."/></xsl:attribute>
						</xsl:element>					
					</xsl:for-each>
				</xsl:for-each>
			</xsl:element>
			<xsl:element name="BESPOKEDECISIONREASONCODESBLOCK">
				<xsl:for-each select="$DEC7/REASONCODEAREA">
					<xsl:element name="BESPOKEDECISIONREASONCODES">
						<xsl:attribute name="SEQUENCENUMBER"><xsl:value-of select="position()"/></xsl:attribute>
						<xsl:attribute name="REASONCODE"><xsl:value-of select="./REASONCODE"/></xsl:attribute>
						<xsl:attribute name="REASONCODEDESCRIPTION"><xsl:value-of select="./CODEDESCRIPTION"/></xsl:attribute>
					</xsl:element>
				</xsl:for-each>
			</xsl:element>
			<xsl:element name="BXCUSTOMERBUREAU1BLOCK">			
				<xsl:element name="BXCUSTOMERBUREAU1">
					<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="$PrimaryApplicant/@CUSTOMERNUMBER"/></xsl:attribute>
					<xsl:attribute name="CUSTOMERVERSIONNUMBER"><xsl:value-of select="$PrimaryApplicant/@CUSTOMERVERSIONNUMBER"/></xsl:attribute>
					<xsl:attribute name="PERSONTYPE"><xsl:value-of select="'1'"/></xsl:attribute>
					<xsl:attribute name="BX01_REFERENCESSCURRENT"><xsl:value-of select="$BX01/E4Q01"/></xsl:attribute>
					<xsl:attribute name="BX01_REFRENCESPCURRENT"><xsl:value-of select="$BX01/E4Q02"/></xsl:attribute>
					<xsl:attribute name="BX01_ERYEARSSSCURRENT"><xsl:value-of select="$BX01/E4Q03"/></xsl:attribute>
					<xsl:attribute name="BX01_ERYEARSSPCURRENT"><xsl:value-of select="$BX01/E4Q04"/></xsl:attribute>
					<xsl:attribute name="BX01_REFERENCESSPREVIOUS"><xsl:value-of select="$BX01/E4Q05"/></xsl:attribute>
					<xsl:attribute name="BX01_REFERENCESPPREVIOUS"><xsl:value-of select="$BX01/E4Q06"/></xsl:attribute>
					<xsl:attribute name="BX01_ERYEARSSSPREVIOUS"><xsl:value-of select="$BX01/E4Q07"/></xsl:attribute>
					<xsl:attribute name="BX01_ERYEARSSPPREVIOUS"><xsl:value-of select="$BX01/E4Q08"/></xsl:attribute>
					<xsl:attribute name="BX01_NAMESCONFIRMED"><xsl:value-of select="$BX01/E4Q17"/></xsl:attribute>
					<xsl:attribute name="BX01_CURRENTADDRESS"><xsl:value-of select="$BX01/E4R01"/></xsl:attribute>
					<xsl:attribute name="BX01_POSTCODE"><xsl:value-of select="$BX01/E4R02"/></xsl:attribute>
					<xsl:attribute name="BX01_PREVIOUSADDRESS"><xsl:value-of select="$BX01/EA4R01PM"/></xsl:attribute>
					<xsl:attribute name="BX01_VOTERSROLLPAF"><xsl:value-of select="$BX01/NDERL01"/></xsl:attribute>
					<xsl:attribute name="BX01_SFREFERENCE"><xsl:value-of select="$BX01/EA2Q01"/></xsl:attribute>
					<xsl:attribute name="BX01_SFYEARSONVOTERSROLL"><xsl:value-of select="$BX01/EA2Q02"/></xsl:attribute>			
					<xsl:attribute name="BX02_TOTALCCJS"><xsl:value-of select="$BX02/E1A01"/></xsl:attribute>			
					<xsl:attribute name="BX02_TOTALVALUECCJS"><xsl:value-of select="$BX02/E1A02"/></xsl:attribute>			
					<xsl:attribute name="BX02_AGEOFLASTCCJ"><xsl:value-of select="$BX02/E1A03"/></xsl:attribute>			
					<xsl:attribute name="BX02_BANKRUPTCYDETECTED"><xsl:value-of select="$BX02/EA1C01"/></xsl:attribute>			
					<xsl:attribute name="BX02_TOTALOUTSTANDINGCCJ"><xsl:value-of select="$BX02/EA1D01"/></xsl:attribute>			
					<xsl:attribute name="BX02_TOTALVALUEOUTSTANDINGCCJ"><xsl:value-of select="$BX02/EA1D02"/></xsl:attribute>			
					<xsl:attribute name="BX02_AGEOFLASTOUTSTANDINGCCJ"><xsl:value-of select="$BX02/EA1D03"/></xsl:attribute>			
					<xsl:attribute name="BX02_BANKRUPTCYPRESENT"><xsl:value-of select="$BX02/SP_BR_PRESENT"/></xsl:attribute>			
					<xsl:attribute name="BX04_TOTALSEARCHESMOCAIS3M"><xsl:value-of select="$BX04/E1E01"/></xsl:attribute>			
					<xsl:attribute name="BX04_TOTALSEARCHESMOCAIS6M"><xsl:value-of select="$BX04/E1E02"/></xsl:attribute>
					<xsl:attribute name="BX04_TOTALOCSEARCHES"><xsl:value-of select="$BX04/EA1B01"/></xsl:attribute>			
					<xsl:attribute name="BX04_TOTALSEARCHESFH3M"><xsl:value-of select="$BX04/NDPSD01"/></xsl:attribute>			
					<xsl:attribute name="BX04_TOTALSEARCHESRETAIL3M"><xsl:value-of select="$BX04/NDPSD02"/></xsl:attribute>			
					<xsl:attribute name="BX04_TOTALSEARCHESBANK3M"><xsl:value-of select="$BX04/NDPSD03"/></xsl:attribute>			
					<xsl:attribute name="BX04_TOTALSEARCHESCC3M"><xsl:value-of select="$BX04/NDPSD04"/></xsl:attribute>			
					<xsl:attribute name="BX04_TOTALSEARCHES6M"><xsl:value-of select="$BX04/NDPSD05"/></xsl:attribute>			
					<xsl:attribute name="BX04_TOTALSEARCHESOC3M"><xsl:value-of select="$BX04/NDPSD06"/></xsl:attribute>			
					<xsl:attribute name="BX04_TOTALSEARCHESNONMO3M"><xsl:value-of select="$BX04/EA1E01"/></xsl:attribute>			
					<xsl:attribute name="BX04_TOTALSEARCHESNONMO6M"><xsl:value-of select="$BX04/EA1E02"/></xsl:attribute>			
					<xsl:attribute name="BX04_TOTALSEARCHESMO3M"><xsl:value-of select="$BX04/EA1E03"/></xsl:attribute>
					<xsl:attribute name="BX04_TOTALSEARCHESMO6M"><xsl:value-of select="$BX04/EA1E04"/></xsl:attribute>			
					<xsl:attribute name="BX05_CIFASDETECTED"><xsl:value-of select="$BX05/EA1A01"/></xsl:attribute>			
					<xsl:attribute name="BX06_CMLDETECTED"><xsl:value-of select="$BX06/EA1C02"/></xsl:attribute>
				</xsl:element>
				<xsl:element name="BXCUSTOMERBUREAU1">
					<xsl:choose>
						<xsl:when test="$SecondApplicant">
							<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="$SecondApplicant/@CUSTOMERNUMBER"/></xsl:attribute>
							<xsl:attribute name="CUSTOMERVERSIONNUMBER"><xsl:value-of select="$SecondApplicant/@CUSTOMERVERSIONNUMBER"/></xsl:attribute>
						</xsl:when>
						<xsl:otherwise>
							<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="$PrimaryApplicant/@CUSTOMERNUMBER"/></xsl:attribute>
							<xsl:attribute name="CUSTOMERVERSIONNUMBER"><xsl:value-of select="$PrimaryApplicant/@CUSTOMERVERSIONNUMBER"/></xsl:attribute>
						</xsl:otherwise>
					</xsl:choose>
						<xsl:attribute name="PERSONTYPE"><xsl:value-of select="'2'"/></xsl:attribute>
						<xsl:attribute name="BX01_REFERENCESSCURRENT"><xsl:value-of select="$BX01/E4Q09"/></xsl:attribute>
						<xsl:attribute name="BX01_REFRENCESPCURRENT"><xsl:value-of select="$BX01/E4Q10"/></xsl:attribute>
						<xsl:attribute name="BX01_ERYEARSSSCURRENT"><xsl:value-of select="$BX01/E4Q11"/></xsl:attribute>
						<xsl:attribute name="BX01_ERYEARSSPCURRENT"><xsl:value-of select="$BX01/E4Q12"/></xsl:attribute>
						<xsl:attribute name="BX01_REFERENCESSPREVIOUS"><xsl:value-of select="$BX01/E4Q13"/></xsl:attribute>
						<xsl:attribute name="BX01_REFERENCESPPREVIOUS"><xsl:value-of select="$BX01/E4Q14"/></xsl:attribute>
						<xsl:attribute name="BX01_ERYEARSSSPREVIOUS"><xsl:value-of select="$BX01/E4Q15"/></xsl:attribute>
						<xsl:attribute name="BX01_ERYEARSSPPREVIOUS"><xsl:value-of select="$BX01/E4Q16"/></xsl:attribute>
						<xsl:attribute name="BX01_NAMESCONFIRMED"><xsl:value-of select="$BX01/E4Q18"/></xsl:attribute>
						<xsl:attribute name="BX01_CURRENTADDRESS"><xsl:value-of select="$BX01/EA4R01CJ"/></xsl:attribute>
						<xsl:attribute name="BX01_POSTCODE"><xsl:value-of select="$BX01/E4R03"/></xsl:attribute>
						<xsl:attribute name="BX01_PREVIOUSADDRESS"><xsl:value-of select="$BX01/EA4R01PJ"/></xsl:attribute>
						<xsl:attribute name="BX01_VOTERSROLLPAF"><xsl:value-of select="$BX01/NDERL02"/></xsl:attribute>
						<xsl:attribute name="BX01_SFREFERENCE"><xsl:value-of select="''"/></xsl:attribute>
						<xsl:attribute name="BX01_SFYEARSONVOTERSROLL"><xsl:value-of select="''"/></xsl:attribute>			
						<xsl:attribute name="BX02_TOTALCCJS"><xsl:value-of select="$BX02/E2G01"/></xsl:attribute>			
						<xsl:attribute name="BX02_TOTALVALUECCJS"><xsl:value-of select="$BX02/E2G02"/></xsl:attribute>			
						<xsl:attribute name="BX02_AGEOFLASTCCJ"><xsl:value-of select="$BX02/E2G03"/></xsl:attribute>			
						<xsl:attribute name="BX02_BANKRUPTCYDETECTED"><xsl:value-of select="$BX02/EA2I01"/></xsl:attribute>			
						<xsl:attribute name="BX02_TOTALOUTSTANDINGCCJ"><xsl:value-of select="$BX02/EA2J01"/></xsl:attribute>			
						<xsl:attribute name="BX02_TOTALVALUEOUTSTANDINGCCJ"><xsl:value-of select="$BX02/EA2J02"/></xsl:attribute>			
						<xsl:attribute name="BX02_AGEOFLASTOUTSTANDINGCCJ"><xsl:value-of select="$BX02/EA2J03"/></xsl:attribute>			
						<xsl:attribute name="BX02_BANKRUPTCYPRESENT"><xsl:value-of select="$BX02/SPA_BR_PRESENT"/></xsl:attribute>			
						<xsl:attribute name="BX04_TOTALSEARCHESMOCAIS3M"><xsl:value-of select="$BX04/E2K01"/></xsl:attribute>			
						<xsl:attribute name="BX04_TOTALSEARCHESMOCAIS6M"><xsl:value-of select="$BX04/E2K02"/></xsl:attribute>
						<xsl:attribute name="BX04_TOTALOCSEARCHES"><xsl:value-of select="$BX04/EA2H01"/></xsl:attribute>			
						<xsl:attribute name="BX04_TOTALSEARCHESFH3M"><xsl:value-of select="$BX04/NDPSD07"/></xsl:attribute>			
						<xsl:attribute name="BX04_TOTALSEARCHESRETAIL3M"><xsl:value-of select="$BX04/NDPSD08"/></xsl:attribute>			
						<xsl:attribute name="BX04_TOTALSEARCHESBANK3M"><xsl:value-of select="$BX04/NDPSD09"/></xsl:attribute>			
						<xsl:attribute name="BX04_TOTALSEARCHESCC3M"><xsl:value-of select="$BX04/NDPSD10"/></xsl:attribute>			
						<xsl:attribute name="BX04_TOTALSEARCHES6M"><xsl:value-of select="''"/></xsl:attribute>			
						<xsl:attribute name="BX04_TOTALSEARCHESOC3M"><xsl:value-of select="''"/></xsl:attribute>			
						<xsl:attribute name="BX04_TOTALSEARCHESNONMO3M"><xsl:value-of select="$BX04/EA2K01"/></xsl:attribute>			
						<xsl:attribute name="BX04_TOTALSEARCHESNONMO6M"><xsl:value-of select="$BX04/EA2K02"/></xsl:attribute>			
						<xsl:attribute name="BX04_TOTALSEARCHESMO3M"><xsl:value-of select="$BX04/EA2K03"/></xsl:attribute>
						<xsl:attribute name="BX04_TOTALSEARCHESMO6M"><xsl:value-of select="$BX04/EA2K04"/></xsl:attribute>			
						<xsl:attribute name="BX05_CIFASDETECTED"><xsl:value-of select="$BX05/EA2G01"/></xsl:attribute>			
						<xsl:attribute name="BX06_CMLDETECTED"><xsl:value-of select="$BX06/EA2I02"/></xsl:attribute>
				</xsl:element>
			</xsl:element>	
			<xsl:element name="BXCUSTOMERBUREAU2BLOCK">
				<xsl:element name="BXCUSTOMERBUREAU2">
					<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="$PrimaryApplicant/@CUSTOMERNUMBER"/></xsl:attribute>
					<xsl:attribute name="CUSTOMERVERSIONNUMBER"><xsl:value-of select="$PrimaryApplicant/@CUSTOMERVERSIONNUMBER"/></xsl:attribute>
					<xsl:attribute name="PERSONTYPE"><xsl:value-of select="1"/></xsl:attribute>
					<xsl:attribute name="BX07_GAINCURRENTADDRESS"><xsl:value-of select="$BX07/EA1G01"/></xsl:attribute>
					<xsl:attribute name="BX07_GAINPREVIOUSADDRESS"><xsl:value-of select="$BX07/EA1G02"/></xsl:attribute>
					<xsl:attribute name="BX11_CONSUMERINDEBTEDNESSINDEX"><xsl:value-of select="$BX11/NDSPCII"/></xsl:attribute>
					<xsl:attribute name="BX12_FAMILYADDRESSPERCENT"><xsl:value-of select="$BX12/EA4M01"/></xsl:attribute>
					<xsl:attribute name="BX12_SINGLESADDRESSPERCENT"><xsl:value-of select="$BX12/EA4M02"/></xsl:attribute>
					<xsl:attribute name="BX12_MIXEDADDRESSPERCENT"><xsl:value-of select="$BX12/EA4M03"/></xsl:attribute>
					<xsl:attribute name="BX12_UNSATISFIEDCCJADDRPERCENT"><xsl:value-of select="$BX12/EA4M04"/></xsl:attribute>
					<xsl:attribute name="BX12_CCJPERELECTORPERCENT"><xsl:value-of select="$BX12/EA4M05"/></xsl:attribute>
					<xsl:attribute name="BX12_DIRECTORPERELECTORPERCENT"><xsl:value-of select="$BX12/EA4M06"/></xsl:attribute>
					<xsl:attribute name="BX12_UNSATISFIEDCCJGT500POUNDS"><xsl:value-of select="$BX12/EA4M07"/></xsl:attribute>
					<xsl:attribute name="BX12_BUILDINGSOCSEARCHPERCENT"><xsl:value-of select="$BX12/EA4M08"/></xsl:attribute>
					<xsl:attribute name="BX12_ELECTRICSEARCHPERCENT"><xsl:value-of select="$BX12/EA4M09"/></xsl:attribute>
					<xsl:attribute name="BX12_FLATYPESEARCHPERCENT"><xsl:value-of select="$BX12/EA4M10"/></xsl:attribute>
					<xsl:attribute name="BX12_PERSONALLOANSEARCHPERCENT"><xsl:value-of select="$BX12/EA4M11"/></xsl:attribute>
					<xsl:attribute name="BX12_UTILITIESSEARCHPERCENT"><xsl:value-of select="$BX12/EA4M12"/></xsl:attribute>
					<xsl:attribute name="BX15_DIRECTORSFLAG"><xsl:value-of select="$BX15/NDDIRSP"/></xsl:attribute>
					<xsl:attribute name="BX17_NDDATEOFBIRTH"><xsl:value-of select="$BX17/NDDOB"/></xsl:attribute>
					<xsl:attribute name="BX17_AGE"><xsl:value-of select="$BX17/EA5S01"/></xsl:attribute>
					<xsl:attribute name="BX18_IMPAIREDCREDITHISTORY"><xsl:value-of select="$BX18/ND_MA_ICH"/></xsl:attribute>
					<xsl:attribute name="BX18_SECUREDARREARS"><xsl:value-of select="$BX18/ND_MA_SEC_ARR"/></xsl:attribute>
					<xsl:attribute name="BX18_UNSECUREDARREARS"><xsl:value-of select="$BX18/ND_MA_UNSEC_ARR"/></xsl:attribute>
					<xsl:attribute name="BX18_CCJAMOUNT"><xsl:value-of select="$BX18/ND_MA_CCJ"/></xsl:attribute>
					<xsl:attribute name="BX18_INDIVIDUALVOLARRANGEMENT"><xsl:value-of select="$BX18/ND_MA_IVA"/></xsl:attribute>
					<xsl:attribute name="BX18_BANKRUPTCY"><xsl:value-of select="$BX18/ND_MA_BANKRUPT"/></xsl:attribute>												
				</xsl:element>
				<xsl:element name="BXCUSTOMERBUREAU2">
					<xsl:choose>
						<xsl:when test="$SecondApplicant">
							<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="$SecondApplicant/@CUSTOMERNUMBER"/></xsl:attribute>
							<xsl:attribute name="CUSTOMERVERSIONNUMBER"><xsl:value-of select="$SecondApplicant/@CUSTOMERVERSIONNUMBER"/></xsl:attribute>
						</xsl:when>
						<xsl:otherwise>
							<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="$PrimaryApplicant/@CUSTOMERNUMBER"/></xsl:attribute>
							<xsl:attribute name="CUSTOMERVERSIONNUMBER"><xsl:value-of select="$PrimaryApplicant/@CUSTOMERVERSIONNUMBER"/></xsl:attribute>
						</xsl:otherwise>
					
						</xsl:choose>
					
					<xsl:attribute name="PERSONTYPE"><xsl:value-of select="2"/></xsl:attribute>
					<xsl:attribute name="BX07_GAINCURRENTADDRESS"><xsl:value-of select="$BX07/EA2M01"/></xsl:attribute>
					<xsl:attribute name="BX07_GAINPREVIOUSADDRESS"><xsl:value-of select="$BX07/EA2M02"/></xsl:attribute>
					<xsl:attribute name="BX11_CONSUMERINDEBTEDNESSINDEX"><xsl:value-of select="$BX11/NDSPACII"/></xsl:attribute>
					<xsl:attribute name="BX12_FAMILYADDRESSPERCENT"><xsl:value-of select="''"/></xsl:attribute>
					<xsl:attribute name="BX12_SINGLESADDRESSPERCENT"><xsl:value-of select="''"/></xsl:attribute>
					<xsl:attribute name="BX12_MIXEDADDRESSPERCENT"><xsl:value-of select="''"/></xsl:attribute>
					<xsl:attribute name="BX12_UNSATISFIEDCCJADDRPERCENT"><xsl:value-of select="''"/></xsl:attribute>
					<xsl:attribute name="BX12_CCJPERELECTORPERCENT"><xsl:value-of select="''"/></xsl:attribute>
					<xsl:attribute name="BX12_DIRECTORPERELECTORPERCENT"><xsl:value-of select="''"/></xsl:attribute>
					<xsl:attribute name="BX12_UNSATISFIEDCCJGT500POUNDS"><xsl:value-of select="''"/></xsl:attribute>
					<xsl:attribute name="BX12_BUILDINGSOCSEARCHPERCENT"><xsl:value-of select="''"/></xsl:attribute>
					<xsl:attribute name="BX12_ELECTRICSEARCHPERCENT"><xsl:value-of select="''"/></xsl:attribute>
					<xsl:attribute name="BX12_FLATYPESEARCHPERCENT"><xsl:value-of select="''"/></xsl:attribute>
					<xsl:attribute name="BX12_PERSONALLOANSEARCHPERCENT"><xsl:value-of select="''"/></xsl:attribute>
					<xsl:attribute name="BX12_UTILITIESSEARCHPERCENT"><xsl:value-of select="''"/></xsl:attribute>
					<xsl:attribute name="BX15_DIRECTORSFLAG"><xsl:value-of select="$BX15/NDDIRSPA"/></xsl:attribute>
					<xsl:attribute name="BX17_NDDATEOFBIRTH"><xsl:value-of select="''"/></xsl:attribute>
					<xsl:attribute name="BX17_AGE"><xsl:value-of select="''"/></xsl:attribute>
					<xsl:attribute name="BX18_IMPAIREDCREDITHISTORY"><xsl:value-of select="$BX18/ND_JA_ICH"/></xsl:attribute>
					<xsl:attribute name="BX18_SECUREDARREARS"><xsl:value-of select="$BX18/ND_JA_SEC_ARR"/></xsl:attribute>
					<xsl:attribute name="BX18_UNSECUREDARREARS"><xsl:value-of select="$BX18/ND_JA_UNSEC_ARR"/></xsl:attribute>
					<xsl:attribute name="BX18_CCJAMOUNT"><xsl:value-of select="$BX18/ND_JA_CCJ"/></xsl:attribute>
					<xsl:attribute name="BX18_INDIVIDUALVOLARRANGEMENT"><xsl:value-of select="$BX18/ND_JA_IVA"/></xsl:attribute>
					<xsl:attribute name="BX18_BANKRUPTCY"><xsl:value-of select="ND_JA_BANKRUPT"/></xsl:attribute>												
				</xsl:element>
			</xsl:element>
			<xsl:element name="BXCUSTOMTERCAISBLOCK">
				<xsl:element name="BXCUSTOMTERCAIS">
					<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="$PrimaryApplicant/@CUSTOMERNUMBER"/></xsl:attribute>
					<xsl:attribute name="CUSTOMERVERSIONNUMBER"><xsl:value-of select="$PrimaryApplicant/@CUSTOMERVERSIONNUMBER"/></xsl:attribute>
					<xsl:attribute name="PERSONTYPE"><xsl:value-of select="'1'"/></xsl:attribute>
					<xsl:attribute name="BX03_TOTALCAIS89"><xsl:value-of select="$BX03/E1A04"/></xsl:attribute>
					<xsl:attribute name="BX03_TOTALVALUECAIS89"><xsl:value-of select="$BX03/E1A05"/></xsl:attribute>
					<xsl:attribute name="BX03_AGEOFLASTCAIS89"><xsl:value-of select="$BX03/E1A06"/></xsl:attribute>
					<xsl:attribute name="BX03_TOTALSETTLEDCAIS"><xsl:value-of select="$BX03/E1A07"/></xsl:attribute>
					<xsl:attribute name="BX03_TOTALSETTLEDCAIS12MONTHS"><xsl:value-of select="$BX03/E1A08"/></xsl:attribute>
					<xsl:attribute name="BX03_TOTALDELINQUENTACC"><xsl:value-of select="$BX03/E1A09"/></xsl:attribute>
					<xsl:attribute name="BX03_TOTALVALUEDELINQUENTACC"><xsl:value-of select="$BX03/E1A10"/></xsl:attribute>
					<xsl:attribute name="BX03_AGEOFLASTDELINQUENTACC"><xsl:value-of select="$BX03/E1A11"/></xsl:attribute>
					<xsl:attribute name="BX03_TOTALACTIVECAIS3M"><xsl:value-of select="$BX03/E1B01"/></xsl:attribute>
					<xsl:attribute name="BX03_TOTALVALUEACTIVECAIS3M"><xsl:value-of select="$BX03/E1B02"/></xsl:attribute>
					<xsl:attribute name="BX03_WORSTSTATUSACTIVECAIS4M"><xsl:value-of select="$BX03/E1B03"/></xsl:attribute>
					<xsl:attribute name="BX03_TOTALVALUEACTIVECAIS4M"><xsl:value-of select="$BX03/E1B04"/></xsl:attribute>
					<xsl:attribute name="BX03_WORSTSTATUSACTIVECAIS12M"><xsl:value-of select="$BX03/E1B05"/></xsl:attribute>
					<xsl:attribute name="BX03_TOTALVALUEACTIVECAIS12M"><xsl:value-of select="$BX03/E1B06"/></xsl:attribute>
					<xsl:attribute name="BX03_WORSTSTATUSALLACTIVECAIS"><xsl:value-of select="$BX03/E1B07"/></xsl:attribute>			
					<xsl:attribute name="BX03_WORSTCURRSTATUSACTIVECAIS"><xsl:value-of select="$BX03/E1B08"/></xsl:attribute>			
					<xsl:attribute name="BX03_TOTALACTIVECAISACCOUNTS"><xsl:value-of select="$BX03/E1B09"/></xsl:attribute>			
					<xsl:attribute name="BX03_TOTALVALACTIVECAISEXCMORT"><xsl:value-of select="$BX03/E1B10"/></xsl:attribute>			
					<xsl:attribute name="BX03_TOTALVALACTIVECAISINCMORT"><xsl:value-of select="$BX03/E1B11"/></xsl:attribute>			
					<xsl:attribute name="BX03_TOTALNOACTIVECAISSTATUS1"><xsl:value-of select="$BX03/E1B12"/></xsl:attribute>			
					<xsl:attribute name="BX03_TOTALNOACTIVECAISSTATUS3"><xsl:value-of select="$BX03/E1B13"/></xsl:attribute>			
					<xsl:attribute name="BX03_TOTALVALUENDCAIS6M"><xsl:value-of select="$BX03/NDECC01"/></xsl:attribute>			
					<xsl:attribute name="BX03_TOTALVALUENDCAISALL"><xsl:value-of select="$BX03/NDECC02"/></xsl:attribute>			
					<xsl:attribute name="BX03_TOTALACTIVENDCAIS6M"><xsl:value-of select="$BX03/NDECC03"/></xsl:attribute>			
					<xsl:attribute name="BX03_RATIOBALANCE3M12M"><xsl:value-of select="$BX03/NDECC04"/></xsl:attribute>
					<xsl:attribute name="BX03_MORTGAGEACCOUNTTIME"><xsl:value-of select="$BX03/NDECC07"/></xsl:attribute>			
					<xsl:attribute name="BX03_MORTGAGEACCOUNTBALANCE"><xsl:value-of select="$BX03/NDECC08"/></xsl:attribute>			
					<xsl:attribute name="BX03_WORSTSTATUSOC"><xsl:value-of select="$BX03/E1C01"/></xsl:attribute>			
					<xsl:attribute name="BX03_TOTALOCACCOUNTS"><xsl:value-of select="$BX03/E1C02"/></xsl:attribute>			
					<xsl:attribute name="BX03_TOTALVALUEOC"><xsl:value-of select="$BX03/E1C03"/></xsl:attribute>			
					<xsl:attribute name="BX03_TOTALDELINQUENTOCACCOUNT"><xsl:value-of select="$BX03/E1C04"/></xsl:attribute>			
					<xsl:attribute name="BX03_TOTALOCCAIS89"><xsl:value-of select="$BX03/E1C05"/></xsl:attribute>			
					<xsl:attribute name="BX03_TOTALSETTLEDOCCAIS"><xsl:value-of select="$BX03/E1C06"/></xsl:attribute>			
					<xsl:attribute name="BX03_TOTALOCCAISACCOUNTS"><xsl:value-of select="$BX03/EA1B02"/></xsl:attribute>			
					<xsl:attribute name="BX03_TOTALMOCAIS89"><xsl:value-of select="$BX03/E1D01"/></xsl:attribute>
					<xsl:attribute name="BX03_TOTALSETTLEDMOCAIS"><xsl:value-of select="$BX03/E1D02"/></xsl:attribute>			
					<xsl:attribute name="BX03_TOTALSETTLEDMOCAIS12M"><xsl:value-of select="$BX03/E1D03"/></xsl:attribute>			
					<xsl:attribute name="BX03_TOTALDELINQUENTMOCAISACC"><xsl:value-of select="$BX03/E1D04"/></xsl:attribute>
					<xsl:attribute name="BX03_WORSTSTATUSNDH3MLT12M"><xsl:value-of select="$BX03/NDHAC01"/></xsl:attribute>
					<xsl:attribute name="BX03_WORSTSTATUSNDH3MGT12M"><xsl:value-of select="$BX03/NDHAC02"/></xsl:attribute>
					<xsl:attribute name="BX03_WORSTSTATUSNDH6M"><xsl:value-of select="$BX03/NDHAC03"/></xsl:attribute>
					<xsl:attribute name="BX03_WORSTSTATUSNDHACTIVE"><xsl:value-of select="$BX03/NDHAC04"/></xsl:attribute>
					<xsl:attribute name="BX03_WORSTSTATUSNDHNONACTIVE"><xsl:value-of select="$BX03/NDHAC05"/></xsl:attribute>
					<xsl:attribute name="BX03_WORSTSTATUSMORTGAGE6M"><xsl:value-of select="$BX03/NDHAC09"/></xsl:attribute>
					<xsl:attribute name="BX03_TOTALNDGOODCAIS"><xsl:value-of select="$BX03/NDINC01"/></xsl:attribute>
					<xsl:attribute name="BX03_MARKETBASEDACTIVITYIND"><xsl:value-of select="$BX03/EA1F01"/></xsl:attribute>
					<xsl:attribute name="BX03_WORSTSTATUSFCS12M"><xsl:value-of select="$BX03/EA1F02"/></xsl:attribute>
					<xsl:attribute name="BX03_WORSTSTATUSTV12M"><xsl:value-of select="$BX03/EA1F03"/></xsl:attribute>
					<xsl:attribute name="BX03_FCSCUGFLAG"><xsl:value-of select="$BX03/NDFCS01"/></xsl:attribute>
					<xsl:attribute name="BX03_CAISSPECIALINSTRUCTIONIND"><xsl:value-of select="$BX03/EA1F04"/></xsl:attribute>
				</xsl:element>
				<xsl:element name="BXCUSTOMTERCAIS">
					<xsl:choose>
						<xsl:when test="$SecondApplicant">
							<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="$SecondApplicant/@CUSTOMERNUMBER"/></xsl:attribute>
							<xsl:attribute name="CUSTOMERVERSIONNUMBER"><xsl:value-of select="$SecondApplicant/@CUSTOMERVERSIONNUMBER"/></xsl:attribute>
						</xsl:when>
						<xsl:otherwise>
							<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="$PrimaryApplicant/@CUSTOMERNUMBER"/></xsl:attribute>
							<xsl:attribute name="CUSTOMERVERSIONNUMBER"><xsl:value-of select="$PrimaryApplicant/@CUSTOMERVERSIONNUMBER"/></xsl:attribute>
						</xsl:otherwise>
					
					</xsl:choose>
					
					<xsl:attribute name="PERSONTYPE"><xsl:value-of select="'2'"/></xsl:attribute>
					<xsl:attribute name="BX03_TOTALCAIS89"><xsl:value-of select="$BX03/E2G04"/></xsl:attribute>
					<xsl:attribute name="BX03_TOTALVALUECAIS89"><xsl:value-of select="$BX03/E2G05"/></xsl:attribute>
					<xsl:attribute name="BX03_AGEOFLASTCAIS89"><xsl:value-of select="$BX03/E2G06"/></xsl:attribute>
					<xsl:attribute name="BX03_TOTALSETTLEDCAIS"><xsl:value-of select="$BX03/E2G07"/></xsl:attribute>
					<xsl:attribute name="BX03_TOTALSETTLEDCAIS12MONTHS"><xsl:value-of select="$BX03/E2G08"/></xsl:attribute>
					<xsl:attribute name="BX03_TOTALDELINQUENTACC"><xsl:value-of select="$BX03/E2G09"/></xsl:attribute>
					<xsl:attribute name="BX03_TOTALVALUEDELINQUENTACC"><xsl:value-of select="$BX03/E2G10"/></xsl:attribute>
					<xsl:attribute name="BX03_AGEOFLASTDELINQUENTACC"><xsl:value-of select="$BX03/E2G11"/></xsl:attribute>
					<xsl:attribute name="BX03_TOTALACTIVECAIS3M"><xsl:value-of select="$BX03/E2H01"/></xsl:attribute>
					<xsl:attribute name="BX03_TOTALVALUEACTIVECAIS3M"><xsl:value-of select="$BX03/E2H02"/></xsl:attribute>
					<xsl:attribute name="BX03_WORSTSTATUSACTIVECAIS4M"><xsl:value-of select="$BX03/E2H03"/></xsl:attribute>
					<xsl:attribute name="BX03_TOTALVALUEACTIVECAIS4M"><xsl:value-of select="$BX03/E2H04"/></xsl:attribute>
					<xsl:attribute name="BX03_WORSTSTATUSACTIVECAIS12M"><xsl:value-of select="$BX03/E2H05"/></xsl:attribute>
					<xsl:attribute name="BX03_TOTALVALUEACTIVECAIS12M"><xsl:value-of select="$BX03/E2H06"/></xsl:attribute>
					<xsl:attribute name="BX03_WORSTSTATUSALLACTIVECAIS"><xsl:value-of select="$BX03/E2H07"/></xsl:attribute>			
					<xsl:attribute name="BX03_WORSTCURRSTATUSACTIVECAIS"><xsl:value-of select="$BX03/E2H08"/></xsl:attribute>			
					<xsl:attribute name="BX03_TOTALACTIVECAISACCOUNTS"><xsl:value-of select="$BX03/E2H09"/></xsl:attribute>			
					<xsl:attribute name="BX03_TOTALVALACTIVECAISEXCMORT"><xsl:value-of select="$BX03/E2H10"/></xsl:attribute>			
					<xsl:attribute name="BX03_TOTALVALACTIVECAISINCMORT"><xsl:value-of select="$BX03/E2H11"/></xsl:attribute>			
					<xsl:attribute name="BX03_TOTALNOACTIVECAISSTATUS1"><xsl:value-of select="$BX03/E2H12"/></xsl:attribute>			
					<xsl:attribute name="BX03_TOTALNOACTIVECAISSTATUS3"><xsl:value-of select="$BX03/E2H13"/></xsl:attribute>			
					<xsl:attribute name="BX03_TOTALVALUENDCAIS6M"><xsl:value-of select="$BX03/NDECC05"/></xsl:attribute>			
					<xsl:attribute name="BX03_TOTALVALUENDCAISALL"><xsl:value-of select="''"/></xsl:attribute>			
					<xsl:attribute name="BX03_TOTALACTIVENDCAIS6M"><xsl:value-of select="''"/></xsl:attribute>			
					<xsl:attribute name="BX03_RATIOBALANCE3M12M"><xsl:value-of select="''"/></xsl:attribute>
					<xsl:attribute name="BX03_MORTGAGEACCOUNTTIME"><xsl:value-of select="$BX03/NDECC09"/></xsl:attribute>			
					<xsl:attribute name="BX03_MORTGAGEACCOUNTBALANCE"><xsl:value-of select="$BX03/NDECC10"/></xsl:attribute>			
					<xsl:attribute name="BX03_WORSTSTATUSOC"><xsl:value-of select="$BX03/E2I01"/></xsl:attribute>			
					<xsl:attribute name="BX03_TOTALOCACCOUNTS"><xsl:value-of select="$BX03/E2I02"/></xsl:attribute>			
					<xsl:attribute name="BX03_TOTALVALUEOC"><xsl:value-of select="$BX03/E2I03"/></xsl:attribute>			
					<xsl:attribute name="BX03_TOTALDELINQUENTOCACCOUNT"><xsl:value-of select="$BX03/E2I04"/></xsl:attribute>			
					<xsl:attribute name="BX03_TOTALOCCAIS89"><xsl:value-of select="$BX03/E2I05"/></xsl:attribute>			
					<xsl:attribute name="BX03_TOTALSETTLEDOCCAIS"><xsl:value-of select="$BX03/E2I06"/></xsl:attribute>			
					<xsl:attribute name="BX03_TOTALOCCAISACCOUNTS"><xsl:value-of select="$BX03/EA2H02"/></xsl:attribute>			
					<xsl:attribute name="BX03_TOTALMOCAIS89"><xsl:value-of select="$BX03/E2J01"/></xsl:attribute>
					<xsl:attribute name="BX03_TOTALSETTLEDMOCAIS"><xsl:value-of select="$BX03/E2J02"/></xsl:attribute>			
					<xsl:attribute name="BX03_TOTALSETTLEDMOCAIS12M"><xsl:value-of select="$BX03/E2J03"/></xsl:attribute>			
					<xsl:attribute name="BX03_TOTALDELINQUENTMOCAISACC"><xsl:value-of select="$BX03/E2J04"/></xsl:attribute>
					<xsl:attribute name="BX03_WORSTSTATUSNDH3MLT12M"><xsl:value-of select="$BX03/NDHAC06"/></xsl:attribute>
					<xsl:attribute name="BX03_WORSTSTATUSNDH3MGT12M"><xsl:value-of select="$BX03/NDHAC07"/></xsl:attribute>
					<xsl:attribute name="BX03_WORSTSTATUSNDH6M"><xsl:value-of select="$BX03/NDHAC08"/></xsl:attribute>
					<xsl:attribute name="BX03_WORSTSTATUSNDHACTIVE"><xsl:value-of select="''"/></xsl:attribute>
					<xsl:attribute name="BX03_WORSTSTATUSNDHNONACTIVE"><xsl:value-of select="''"/></xsl:attribute>
					<xsl:attribute name="BX03_WORSTSTATUSMORTGAGE6M"><xsl:value-of select="$BX03/NDHAC10"/></xsl:attribute>
					<xsl:attribute name="BX03_TOTALNDGOODCAIS"><xsl:value-of select="$BX03/NDINC02"/></xsl:attribute>
					<xsl:attribute name="BX03_MARKETBASEDACTIVITYIND"><xsl:value-of select="$BX03/EA2L01"/></xsl:attribute>
					<xsl:attribute name="BX03_WORSTSTATUSFCS12M"><xsl:value-of select="$BX03/EA2L02"/></xsl:attribute>
					<xsl:attribute name="BX03_WORSTSTATUSTV12M"><xsl:value-of select="$BX03/EA2L03"/></xsl:attribute>
					<xsl:attribute name="BX03_FCSCUGFLAG"><xsl:value-of select="$BX03/NDFCS02"/></xsl:attribute>
					<xsl:attribute name="BX03_CAISSPECIALINSTRUCTIONIND"><xsl:value-of select="$BX03/EA2L04"/></xsl:attribute>
				</xsl:element>
			</xsl:element>
			<xsl:element name="BESPOKEBUREAUDATABLOCK">
				<xsl:element name="BESPOKEBUREAUDATA">
					<xsl:attribute name="BES1_NUMBERUNSATISFIEDCCJ"><xsl:value-of select="$BES1/NO_UNSAT_CCJ"/></xsl:attribute>					
					<xsl:attribute name="BES1_NUMBERSATISFIEDCCJ"><xsl:value-of select="$BES1/NO_SAT_CCJ"/></xsl:attribute>					
					<xsl:attribute name="BES1_VALUEUNSATISFIEDCCJ"><xsl:value-of select="$BES1/VAL_UNSAT_CCJ"/></xsl:attribute>					
					<xsl:attribute name="BES1_VALUESATISFIEDCCJ"><xsl:value-of select="$BES1/VAL_SAT_CCJ"/></xsl:attribute>					
					<xsl:attribute name="BES1_MOSTRECENTREGDATEUNSATCCJ"><xsl:value-of select="$BES1/MREC_RDATE_UNSAT_CCJ"/></xsl:attribute>					
					<xsl:attribute name="BES1_MOSTRECENTREGDATESATCCJ"><xsl:value-of select="$BES1/MREC_RDATE_SAT_CCJ"/></xsl:attribute>					
					<xsl:attribute name="BES1_MOSTRECENTSATDATESATCJJ"><xsl:value-of select="$BES1/MREC_SDATE_SAT_CCJ"/></xsl:attribute>					
					<xsl:attribute name="BES1_MOSTRECENTBANKRUPTCYDISDATE"><xsl:value-of select="$BES1/MREC_DDATE_BNKRUP"/></xsl:attribute>					
					<xsl:attribute name="BES1_SECUREDRENTARREARSWRSTSTATL12M"><xsl:value-of select="$BES1/SECRENT_WS_12M"/></xsl:attribute>					
					<xsl:attribute name="BES1_DATELATESTBANKRUPTCY"><xsl:value-of select="$BES1/MREC_BNKRUP_DATE"/></xsl:attribute>					
					<xsl:attribute name="BES1_IVAWORSTSTATUS"><xsl:value-of select="$BES1/IVA_WST_STATE"/></xsl:attribute>					
					<xsl:attribute name="BES1_IVADATE"><xsl:value-of select="$BES1/IVA_DATE"/></xsl:attribute>					
					<xsl:attribute name="BES1_DATELATESTREPOSSESSION"><xsl:value-of select="$BES1/DATE_LTST_REPOS"/></xsl:attribute>
					<xsl:attribute name="BES1_VALUEUNSATISFIEDCCJL24M"><xsl:value-of select="$BES1/VAL_UNSAT_CCJ_L24M"/></xsl:attribute>
					<xsl:attribute name="BES1_VALUESATISFIEDCCJL12M"><xsl:value-of select="$BES1/VAL_SAT_CCJ_L12M"/></xsl:attribute>
					<xsl:attribute name="BES1_ACTIVEBANKRUPTCY"><xsl:value-of select="$BES1/ACT_BNKRPT"/></xsl:attribute>
					<xsl:attribute name="BES1_VALUECAIS89L24M"><xsl:value-of select="$BES1/VAL_CAIS89_L24M"/></xsl:attribute>
					<xsl:attribute name="BES1_SECURED_RENT_MISSED_L12M"><xsl:value-of select="$BES1/SECRENT_MISS_L12M"/></xsl:attribute>
					<xsl:attribute name="BES1_SECURED_RENT_MISSED_L6M"><xsl:value-of select="$BES1/SECRENT_MISS_L6M"/></xsl:attribute>
					<xsl:attribute name="BES1_SECURED_RENT_WCS"><xsl:value-of select="$BES1/SECRENT_WCS"/></xsl:attribute>
					<xsl:attribute name="BES1_SECURED_RENT_MISSED_L3M"><xsl:value-of select="$BES1/SECRENT_MISS_L3M"/></xsl:attribute>
					<xsl:attribute name="BES1_SECURED_RENT_MISSED_L12M_DEDUPE"><xsl:value-of select="$BES1/SECRENT_MISS_L12M_DE_DUPE"/></xsl:attribute>
					<xsl:attribute name="BES1_SECURED_RENT_MISSED_L6M_DEDUPE"><xsl:value-of select="$BES1/SECRENT_MISS_L6M_DE_DUPE"/></xsl:attribute>
					<xsl:attribute name="BES1_SECURED_RENT_MISSED_L3M_DEDUPE"><xsl:value-of select="$BES1/SECRENT_MISS_L3M_DE_DUPE"/></xsl:attribute>
					<xsl:attribute name="BES1_VALUESATISFIEDCCJ_L6M"><xsl:value-of select="$BES1/VAL_SATISFIED_CCJ_L6M"/></xsl:attribute>
				</xsl:element>
			</xsl:element>
		</xsl:element>
	</xsl:template>
	<xsl:template name="WriteAddressRequiringUpdate">
		<xsl:param name="Seq"/>
		<xsl:for-each select="$AddressMap/ADDRESS[@BLOCKSEQNUMBER = $Seq]">
			<xsl:call-template name="WriteAddressLine">
				<xsl:with-param name="AddressMapLine" select="."/>
				<xsl:with-param name="AddressDetailLine" select="$TargetingData/ADDRESSTARGET[@BLOCKSEQNUMBER = $Seq and @ADDRESSRESOLVED = 'Y']"/>
			</xsl:call-template>			
		</xsl:for-each>
	</xsl:template>
	<xsl:template name="WriteAddressLine">
		<xsl:param name="AddressMapLine"/>
		<xsl:param name="AddressDetailLine"/>
		<xsl:if test="$AddressDetailLine">		
			<xsl:element name="ADDRESS">
				<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="$AddressMapLine/@CUSTOMERNUMBER"/></xsl:attribute>
				<xsl:attribute name="CUSTOMERVERSIONNUMBER"><xsl:value-of select="$AddressMapLine/@CUSTOMERVERSIONNUMBER"/></xsl:attribute>
				<xsl:attribute name="TARGETADDRESSTYPE"><xsl:value-of select="$AddressMapLine/@TARGETADDRESSTYPE"/></xsl:attribute>
				<xsl:attribute name="DATEMOVEDIN"><xsl:value-of select="$AddressMapLine/@DATEMOVEDIN"/></xsl:attribute>
				<xsl:attribute name="DATEMOVEDOUT"><xsl:value-of select="$AddressMapLine/@DATEMOVEDOUT"/></xsl:attribute>
				<xsl:attribute name="NATUREOFOCCUPANCY"><xsl:value-of select="$AddressMapLine/@NATUREOFOCCUPANCY"/></xsl:attribute>
				<xsl:attribute name="BUILDINGORHOUSENAME"><xsl:value-of select="$AddressDetailLine/@HOUSENAME"/></xsl:attribute>
				<xsl:attribute name="BUILDINGORHOUSENUMBER"><xsl:value-of select="$AddressDetailLine/@HOUSENUMBER"/></xsl:attribute>					
				<xsl:attribute name="FLATNUMBER"><xsl:value-of select="$AddressDetailLine/@FLAT"/></xsl:attribute>
				<xsl:attribute name="STREET"><xsl:value-of select="$AddressDetailLine/@STREET"/></xsl:attribute>
				<xsl:attribute name="DISTRICT"><xsl:value-of select="$AddressDetailLine/@DISTRICT"/></xsl:attribute>
				<xsl:attribute name="TOWN"><xsl:value-of select="$AddressDetailLine/@TOWN"/></xsl:attribute>
				<xsl:attribute name="COUNTY"><xsl:value-of select="$AddressDetailLine/@COUNTY"/></xsl:attribute>
				<xsl:attribute name="POSTCODE"><xsl:value-of select="$AddressDetailLine/@POSTCODE"/></xsl:attribute>					
			</xsl:element>
		</xsl:if>	
	</xsl:template>
</xsl:stylesheet>