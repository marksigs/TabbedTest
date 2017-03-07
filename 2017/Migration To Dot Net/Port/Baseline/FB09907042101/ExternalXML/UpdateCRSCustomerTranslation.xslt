<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>

	<xsl:template match="node()|@*">
		<xsl:copy>
			<xsl:apply-templates select="node()|@*"/>
		</xsl:copy>
	</xsl:template>
	
	
	<xsl:template match="AddressType">
		<xsl:element name="{name()}">
			<xsl:choose>
				<xsl:when test=".='H'">H</xsl:when>
				<xsl:when test=".='C'">M</xsl:when>
				<xsl:when test=".='P'">P</xsl:when>
			</xsl:choose>
		</xsl:element>
	</xsl:template>
	
	<xsl:template match="CountryCode">
		<xsl:element name="{name()}">
			<xsl:choose>
				<xsl:when test=".='1'">UK</xsl:when>
				<xsl:when test=".=''">UK</xsl:when>
			</xsl:choose>
		</xsl:element>
	</xsl:template>

	<xsl:template match="PhoneType">
		<xsl:element name="{name()}">
			<xsl:choose>
				<xsl:when test=".='1'">R</xsl:when>
				<xsl:when test=".='2'">B</xsl:when>
				<xsl:when test=".='3'">C</xsl:when>
				<xsl:when test=".='4'">F</xsl:when>
				<xsl:when test=".='5'">C</xsl:when>
			</xsl:choose>
		</xsl:element>
	</xsl:template>
	
	<xsl:template match="Gender">
		<xsl:element name="{name()}">
			<xsl:choose>
				<xsl:when test=".='1'">M</xsl:when>
				<xsl:when test=".='2'">F</xsl:when>
				<xsl:otherwise>X</xsl:otherwise>
			</xsl:choose>
		</xsl:element>
	</xsl:template>
	
	<xsl:template match="SpecialNeed">
		<xsl:element name="{name()}">
			<xsl:choose>
				<xsl:when test=".='2'">L</xsl:when>
				<xsl:when test=".='3'">A</xsl:when>
				<xsl:when test=".='4'">B</xsl:when>
				<xsl:otherwise>0</xsl:otherwise>
			</xsl:choose>
		</xsl:element>
	</xsl:template>
	
		<xsl:template match="DontSolicit">
		<xsl:element name="{name()}">
			<xsl:choose>
				<xsl:when test=".='YES'">false</xsl:when>
				<xsl:when test=".='NO'">true</xsl:when>
			</xsl:choose>
		</xsl:element>
	</xsl:template>


	<xsl:template match="NamePrefix">
		<xsl:element name="{name()}">
			<xsl:choose>
				<xsl:when test=".='1'">Mr.</xsl:when>
				<xsl:when test=".='2'">Mrs.</xsl:when>
				<xsl:when test=".='3'">Miss</xsl:when>
				<xsl:when test=".='4'">Ms</xsl:when>
				<xsl:when test=".='5'">Dr.</xsl:when>
				<xsl:when test=".='53'">Rev.</xsl:when>
				<xsl:when test=".='6'">Adml.</xsl:when>
				<xsl:when test=".='7'">Air_Cm</xsl:when>
				<xsl:when test=".='8'">Air_M.</xsl:when>
				<xsl:when test=".='9'">Air_Vm.</xsl:when>
				<xsl:when test=".='52'">R/Adml</xsl:when>
				<xsl:when test=".='51'">Prof</xsl:when>
				<xsl:when test=".='50'">Private</xsl:when>
				<xsl:when test=".='49'">Princess</xsl:when>
				<xsl:when test=".='48'">Prince</xsl:when>
				<xsl:when test=".='47'">Plt_Off</xsl:when>
				<xsl:when test=".='46'">Mraf</xsl:when>
				<xsl:when test=".='45'">Master</xsl:when>
				<xsl:when test=".='44'">Marquis</xsl:when>
				<xsl:when test=".='43'">Marquess</xsl:when>
				<xsl:when test=".='42'">Major</xsl:when>
				<xsl:when test=".='41'">Madam</xsl:when>
				<xsl:when test=".='40'">Lt_Col.</xsl:when>
				<xsl:when test=".='39'">Lt.</xsl:when>
				<xsl:when test=".='38'">Lord</xsl:when>
				<xsl:when test=".='37'">Lady</xsl:when>
				<xsl:when test=".='36'">L/Cpl.</xsl:when>
				<xsl:when test=".='35'">Judge</xsl:when>
				<xsl:when test=".='34'">HRH</xsl:when>
				<xsl:when test=".='33'">Hon_Mrs.</xsl:when>
				<xsl:when test=".='32'">Hon_Mr.</xsl:when>
				<xsl:when test=".='31'">Grp_Capt.</xsl:when>
				<xsl:when test=".='30'">Gen.</xsl:when>
				<xsl:when test=".='29'">Flt_Lt.</xsl:when>
				<xsl:when test=".='28'">Fg_Off.</xsl:when>
				<xsl:when test=".='27'">Father</xsl:when>
				<xsl:when test=".='26'">F/Sgt.</xsl:when>
				<xsl:when test=".='25'">Earl</xsl:when>
				<xsl:when test=".='24'">Duke</xsl:when>
				<xsl:when test=".='23'">Duchess</xsl:when>
				<xsl:when test=".='22'">Dame</xsl:when>
				<xsl:when test=".='21'">Cpl.</xsl:when>
				<xsl:when test=".='20'">Countess</xsl:when>
				<xsl:when test=".='19'">Count</xsl:when>
				<xsl:when test=".='18'">Col.</xsl:when>
				<xsl:when test=".='17'">Cmdr.</xsl:when>
				<xsl:when test=".='16'">Cdr.</xsl:when>
				<xsl:when test=".='15'">Capt.</xsl:when>
				<xsl:when test=".='14'">Canon</xsl:when>
				<xsl:when test=".='13'">C/Sgt.</xsl:when>
				<xsl:when test=".='12'">Brig.</xsl:when>
				<xsl:when test=".='11'">Baroness</xsl:when>
				<xsl:when test=".='10'">Baron</xsl:when>
				<xsl:when test=".='15'">Wo.</xsl:when>
				<xsl:when test=".='68'">Wg._Cmdr.</xsl:when>
				<xsl:when test=".='67'">Viscount</xsl:when>
				<xsl:when test=".='66'">V/Adml.</xsl:when>
				<xsl:when test=".='65'">The_Rt_Hon</xsl:when>
				<xsl:when test=".='64'">The_Hon</xsl:when>
				<xsl:when test=".='63'">Sqn._Ldr.</xsl:when>
				<xsl:when test=".='62'">Sister</xsl:when>
				<xsl:when test=".='61'">Sir</xsl:when>
				<xsl:when test=".='60'">Sheriff</xsl:when>
				<xsl:when test=".='59'">Sheikh</xsl:when>
				<xsl:when test=".='58'">Sgt.</xsl:when>
				<xsl:when test=".='57'">Sgn._Capt.</xsl:when>
				<xsl:when test=".='56'">Sac.</xsl:when>
				<xsl:when test=".='55'">S/Sgt.</xsl:when>
				<xsl:when test=".='15'">Capt</xsl:when>
				<xsl:when test=".='54'">Rev._Dr.</xsl:when>
				<xsl:when test=".='99'">Other</xsl:when>
			</xsl:choose>
		</xsl:element>
	</xsl:template>
	
</xsl:stylesheet>
