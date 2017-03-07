<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	
	<xsl:template match="/">
	
		<xsl:variable name="AlphaPlus" select="//ALPHAPLUS"/>
		<xsl:element name="RESPONSE">
			<xsl:attribute name="TYPE">SUCCESS</xsl:attribute>
	
			<xsl:element name="APPLICATION">
				<xsl:element name="APPLICATIONFACTFIND">
					<xsl:element name="QUOTATION">
						<xsl:element name="MORTGAGESUBQUOTE">
							<xsl:attribute name="TOTALGROSSMONTHLYCOST"><xsl:value-of select="$AlphaPlus/TOTALGROSSMONTHLYCOST"/></xsl:attribute>
							<xsl:if test="string($AlphaPlus/MONTHLYCOSTLESSDRAWDOWN)">
								<xsl:attribute name="MONTHLYCOSTLESSDRAWDOWN"><xsl:value-of select="$AlphaPlus/MONTHLYCOSTLESSDRAWDOWN"/></xsl:attribute>
							</xsl:if>
							<xsl:attribute name="TOTALAMOUNTPAYABLE"><xsl:value-of select="$AlphaPlus/TOTALAMOUNTPAYABLE"/></xsl:attribute>
							<xsl:attribute name="TOTALMORTGAGEPAYMENTS"><xsl:value-of select="$AlphaPlus/TOTALMORTGAGEPAYMENTS"/></xsl:attribute>
							<xsl:attribute name="TOTALACCRUEDINTEREST"><xsl:value-of select="$AlphaPlus/TOTALACCRUEDINTEREST"/></xsl:attribute>
							<xsl:attribute name="AMOUNTPERUNITBORROWED"><xsl:value-of select="$AlphaPlus/AMOUNTPERUNITBORROWED"/></xsl:attribute>
							<xsl:attribute name="APR"><xsl:value-of select="$AlphaPlus/APR"/></xsl:attribute>
							<xsl:element name="LOANCOMPONENTLIST">
								<xsl:for-each select="$AlphaPlus/LOANCOMPONENTLIST/LOANCOMPONENT">
									<xsl:variable name="LCSeqNum" select="LOANCOMPONENTSEQUENCENUMBER"/>
									<xsl:variable name="LC" select="//CALCS/APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE/LOANCOMPONENTLIST/LOANCOMPONENT[position()=$LCSeqNum]"/>
									<xsl:element name="LOANCOMPONENT">
										<xsl:attribute name="MORTGAGEPRODUCTCODE"><xsl:value-of select="$LC/MORTGAGEPRODUCTCODE"/></xsl:attribute>
										<xsl:attribute name="FINALRATEMONTHLYCOST"><xsl:value-of select="FINALRATEMONTHLYCOST"/></xsl:attribute>
										<!--<xsl:attribute name="PURPOSEOFLOAN"></xsl:attribute>-->
										<xsl:attribute name="ACCRUEDINTEREST"><xsl:value-of select="ACCRUEDINTEREST"/></xsl:attribute>
										<xsl:attribute name="TOTALAMOUNTPAYABLE"><xsl:value-of select="TOTALAMOUNTPAYABLE"/></xsl:attribute>
										<xsl:if test="string(TOTALLOANCOMPONENTAMOUNT)">
											<xsl:attribute name="TOTALLOANCOMPONENTAMOUNT"><xsl:value-of select="TOTALLOANCOMPONENTAMOUNT"/></xsl:attribute>
										</xsl:if>
										<xsl:if test="string(PORTEDLOAN)">
											<xsl:attribute name="PORTEDLOAN"><xsl:value-of select="PORTEDLOAN"/></xsl:attribute>
										</xsl:if>
										<xsl:attribute name="LOANCOMPONENTSEQUENCENUMBER"><xsl:value-of select="$LCSeqNum"/></xsl:attribute>
										<xsl:attribute name="NETINTONLYELEMENT">
											<xsl:choose>
												<xsl:when test="$LC/NETINTONLYELEMENT"><xsl:value-of select="$LC/NETINTONLYELEMENT"/></xsl:when>
												<xsl:otherwise>0</xsl:otherwise>
											</xsl:choose>
										</xsl:attribute>
										<xsl:attribute name="INTERESTONLYELEMENT">
											<xsl:choose>
												<xsl:when test="$LC/INTERESTONLYELEMENT"><xsl:value-of select="$LC/INTERESTONLYELEMENT"/></xsl:when>
												<xsl:otherwise>0</xsl:otherwise>
											</xsl:choose>
										</xsl:attribute>
										<xsl:attribute name="TERMINMONTHS"><xsl:value-of select="$LC/TERMINMONTHS"/></xsl:attribute>
										<xsl:attribute name="TERMINYEARS"><xsl:value-of select="$LC/TERMINYEARS"/></xsl:attribute>
										<xsl:attribute name="GROSSMONTHLYCOST"><xsl:value-of select="GROSSMONTHLYCOST"/></xsl:attribute>
										<xsl:attribute name="NETMONTHLYCOST"><xsl:value-of select="NETMONTHLYCOST"/></xsl:attribute>
										<xsl:if test="string(MONTHLYCOSTLESSDRAWDOWN)">
											<xsl:attribute name="MONTHLYCOSTLESSDRAWDOWN"><xsl:value-of select="MONTHLYCOSTLESSDRAWDOWN"/></xsl:attribute>
										</xsl:if>
										<xsl:attribute name="AMOUNTPERUNITBORROWED"><xsl:value-of select="AMOUNTPERUNITBORROWED"/></xsl:attribute>
										<xsl:attribute name="FINALPAYMENT"><xsl:value-of select="FINALPAYMENT"/></xsl:attribute>
										<xsl:attribute name="NETCAPANDINTELEMENT">
											<xsl:choose>
												<xsl:when test="$LC/NETCAPANDINTELEMENT"><xsl:value-of select="$LC/NETCAPANDINTELEMENT"/></xsl:when>
												<xsl:otherwise>0</xsl:otherwise>
											</xsl:choose>
										</xsl:attribute>
										<xsl:attribute name="CAPITALANDINTERESTELEMENT">
											<xsl:choose>
												<xsl:when test="$LC/CAPITALANDINTERESTELEMENT"><xsl:value-of select="$LC/CAPITALANDINTERESTELEMENT"/></xsl:when>
												<xsl:otherwise>0</xsl:otherwise>
											</xsl:choose>
										</xsl:attribute>
										<xsl:attribute name="FINALRATE"><xsl:value-of select="FINALRATE"/></xsl:attribute>
										<xsl:attribute name="RESOLVEDRATE"><xsl:value-of select="RESOLVEDRATE"/></xsl:attribute>
										<xsl:attribute name="APR"><xsl:value-of select="APR"/></xsl:attribute>
										<xsl:attribute name="REPAYMENTMETHOD"><xsl:value-of select="$LC/REPAYMENTMETHOD"/></xsl:attribute>
										<xsl:for-each select="$AlphaPlus/LOANCOMPONENTLIST/LOANCOMPONENTPAYMENTSCHEDULE[LOANCOMPONENTSEQUENCENUMBER=$LCSeqNum]">
											<xsl:element name="LOANCOMPONENTPAYMENTSCHEDULE">
												<!--<xsl:attribute name="PAYMENTTYPE"><xsl:value-of select=""/></xsl:attribute>-->
												<xsl:attribute name="STARTDATE"><xsl:value-of select="STARTDATE"/></xsl:attribute>
												<xsl:if test="string(MONTHLYCOST)">
													<xsl:attribute name="MONTHLYCOST"><xsl:value-of select="MONTHLYCOST"/></xsl:attribute>
												</xsl:if>
												<xsl:if test="string(MAXMONTHLYCOST)">
													<xsl:attribute name="MAXMONTHLYCOST"><xsl:value-of select="MAXMONTHLYCOST"/></xsl:attribute>
												</xsl:if>
												<xsl:if test="string(MINMONTHLYCOST)">
													<xsl:attribute name="MINMONTHLYCOST"><xsl:value-of select="MINMONTHLYCOST"/></xsl:attribute>
												</xsl:if>
												<xsl:if test="string(INCREASEDMONTHLYCOST)">
													<xsl:attribute name="INCREASEDMONTHLYCOST"><xsl:value-of select="INCREASEDMONTHLYCOST"/></xsl:attribute>
												</xsl:if>
												<xsl:attribute name="LOANCOMPONENTSEQUENCENUMBER"><xsl:value-of select="$LCSeqNum"/></xsl:attribute>
												<xsl:if test="string(INTONLYMONTHLYCOST)">
													<xsl:attribute name="INTONLYMONTHLYCOST"><xsl:value-of select="INTONLYMONTHLYCOST"/></xsl:attribute>
												</xsl:if>												
												<xsl:attribute name="INTERESTRATETYPESEQUENCENUMBER"><xsl:value-of select="INTERESTRATETYPESEQUENCENUMBER"/></xsl:attribute>
												<xsl:if test="INCREASEDMONTHLYCOSTDIFFERENCE">
													<xsl:attribute name="INCREASEDMONTHLYCOSTDIFFERENCE"><xsl:value-of select="INCREASEDMONTHLYCOSTDIFFERENCE"/></xsl:attribute>
												</xsl:if>
												<xsl:if test="string(CAPINTMONTHLYCOST)">
													<xsl:attribute name="CAPINTMONTHLYCOST"><xsl:value-of select="CAPINTMONTHLYCOST"/></xsl:attribute>
												</xsl:if>
												<xsl:attribute name="INTERESTRATE"><xsl:value-of select="INTERESTRATE"/></xsl:attribute>
											</xsl:element>
										</xsl:for-each>
										<xsl:for-each select="$AlphaPlus/LOANCOMPONENTLIST/LOANCOMPONENTREDEMPTIONFEE[LOANCOMPONENTSEQUENCENUMBER=$LCSeqNum]">
											<xsl:element name="LOANCOMPONENTREDEMPTIONFEE">
												<!--<xsl:attribute name="MORTGAGESUBQUOTENUMBER"><xsl:value-of select="MORTGAGESUBQUOTENUMBER"/></xsl:attribute>-->
												<xsl:attribute name="REDEMPTIONFEEAMOUNT"><xsl:value-of select="REDEMPTIONFEEAMOUNT"/></xsl:attribute>
												<xsl:attribute name="REDEMPTIONFEESTEPNUMBER"><xsl:value-of select="REDEMPTIONFEESTEPNUMBER"/></xsl:attribute>
												<xsl:if test="REDEMPTIONFEEPERIOD">
													<xsl:attribute name="REDEMPTIONFEEPERIOD"><xsl:value-of select="REDEMPTIONFEEPERIOD"/></xsl:attribute>
												</xsl:if>
												<xsl:attribute name="LOANCOMPONENTSEQUENCENUMBER"><xsl:value-of select="$LCSeqNum"/></xsl:attribute>
												<!--<xsl:attribute name="APPLICATIONNUMBER"><xsl:value-of select="APPLICATIONNUMBER"/></xsl:attribute>-->
												<xsl:attribute name="REDEMPTIONFEEPERIODENDDATE"><xsl:value-of select="REDEMPTIONFEEPERIODENDDATE"/></xsl:attribute>
												<!--<xsl:attribute name="APPLICATIONFACTFINDNUMBER"><xsl:value-of select="APPLICATIONFACTFINDNUMBER"/></xsl:attribute>-->
											</xsl:element>
										</xsl:for-each>
										<xsl:for-each select="$AlphaPlus/LOANCOMPONENTLIST/LOANCOMPONENTBALANCESCHEDULE[LOANCOMPONENTSEQUENCENUMBER=$LCSeqNum]">
											<xsl:element name="LOANCOMPONENTBALANCESCHEDULE">
												<xsl:attribute name="SCHEDULETYPE"><xsl:value-of select="SCHEDULETYPE"/></xsl:attribute>
												<xsl:attribute name="STARTDATE"><xsl:value-of select="STARTDATE"/></xsl:attribute>
												<xsl:attribute name="INTONLYBALANCE"><xsl:value-of select="INTONLYBALANCE"/></xsl:attribute>
												<xsl:attribute name="LOANCOMPONENTSEQUENCENUMBER"><xsl:value-of select="$LCSeqNum"/></xsl:attribute>
												<xsl:attribute name="BALANCE"><xsl:value-of select="BALANCE"/></xsl:attribute>
												<xsl:attribute name="CAPINTBALANCE"><xsl:value-of select="CAPINTBALANCE"/></xsl:attribute>
											</xsl:element>
										</xsl:for-each>
									</xsl:element>
								</xsl:for-each>
							</xsl:element>
						</xsl:element>
					</xsl:element>
				</xsl:element>
				<xsl:for-each select="ONEOFFCOSTLIST/ONEOFFCOST[number(@AMOUNT)>0]">
                                        <xsl:element name="ONEOFFCOST">
					        <xsl:copy-of select="@*[name()!='IDENTIFIER']"/>
				        </xsl:element>
				</xsl:for-each>
			</xsl:element>
			<xsl:for-each select="//CUTOMERLIST/CUSTOMER">
				<xsl:element name="CUSTOMER">
					<xsl:element name="CUSTOMERVERSION">
						<xsl:attribute name="TITLE"><xsl:value-of select="@TITLE"/></xsl:attribute>
						<xsl:attribute name="OTHERFORENAMES"><xsl:value-of select="@OTHERNAMES"/></xsl:attribute>
						<xsl:attribute name="FIRSTFORENAME"><xsl:value-of select="@FIRSTNAME"/></xsl:attribute>
						<!--<xsl:attribute name="SECONDFORENAME"></xsl:attribute>-->
						<!--<xsl:attribute name="TITLEOTHER"></xsl:attribute>-->
						<xsl:attribute name="SURNAME"><xsl:value-of select="@SURNAME"/></xsl:attribute>
					</xsl:element>
				</xsl:element>
			</xsl:for-each>
		</xsl:element>
	</xsl:template>
</xsl:stylesheet>
