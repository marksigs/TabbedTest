 <xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">

	<xsl:variable name="omigaEntities" select="/REQUEST/ENTITIES" />
	<xsl:variable name="exclusions" select="/REQUEST/EXCLUSIONS" />
	
	<xsl:template match="/">
		<xsl:element name="REQUESTType">
			
			<xsl:choose>
				<xsl:when test="REQUEST/APPLICATION/@TYPE='SUBMITFMA'">
					<xsl:element name="APPLICATION" namespace="http://Request.SubmitFMA.Omiga.vertex.co.uk">
						<xsl:apply-templates select="REQUEST/APPLICATION" />					
					</xsl:element>
				</xsl:when>
				<xsl:when test="REQUEST/APPLICATION/@TYPE='STOPANDSAVE'">
					<xsl:element name="APPLICATION" namespace="http://Request.SubmitStopAndSaveFMA.Omiga.vertex.co.uk">
						<xsl:apply-templates select="REQUEST/APPLICATION" />
					</xsl:element>					
				</xsl:when>
			</xsl:choose>
		</xsl:element>
	</xsl:template>

	<!-- APPLICATION -->
	<xsl:template match="/REQUEST/APPLICATION">
		<xsl:variable name="omigaAppRoot" select="//REQUEST/APPLICATIONType" />

		<!-- Output attributes that exist on context omiga node only -->
		<!--<xsl:apply-templates select="$omigaAppRoot/@*" mode="OmigaOnlyAttributes">
		<xsl:with-param name="webContextNode" select="." />
		</xsl:apply-templates>-->

		<!-- Copy all web context attributes. They're either identical or changed from the omiga values and dont exist in the output --> 
		<xsl:if test="count(@*)>0">			
			<xsl:copy-of select="@*[name()!='TYPE']" /><!-- Exclude: TYPE att. which is to establish namespace only-->
			<xsl:attribute name="CRUD_OP">			
				<xsl:value-of select="'UPDATE'" />
			</xsl:attribute>
		</xsl:if>
			
		<!-- Process APPLICATION child Elements -->
		<xsl:apply-templates select="*">
			<xsl:with-param name="omigaPath" select="$omigaAppRoot/*"/>
		</xsl:apply-templates>
	</xsl:template>
	<!-- /APPLICATION -->

	<!-- ELEMENTS -->
	<xsl:template match="*">

		<!-- Param
			$omigaPath: Is a reference to the parent of the context node in the associated omiga data i.e:
				context 	   =	REQUEST/APPLICATION[position()=1]/APPLICATIONFACTFIND
				$omigaPath  =	REQUEST/APPLICATION[position()=2]/
		-->
		<xsl:param name="omigaPath"/>
		
		<!-- Variables
			Establish $omigaNode:
			$omigaNode is calculated by looking for the child node under $omigaPath that has the same name as the context 'web data' node and the 
			same key data.
		-->		
		<xsl:variable name="contextNodeName" select="local-name()" />		<!-- Get string name of the context node. Used to get Key data in contextKey -->		
		<xsl:variable name="contextKey" select="$omigaEntities/ENTITY[@NAME=$contextNodeName]/KEY/@NAME" />	<!-- Get the Primary key attribute names for the context node. Defined in OmigaEntityKeys.xml -->		
		<xsl:variable name="contextKeyValues" select="@*[name()=$contextKey]" />	<!-- Get Nodeset of the actual context nodes key data as defined in $contextKey-->

		<xsl:variable name="omigaContextPosition">
			<xsl:choose>
				<xsl:when test="count($omigaPath[name()=$contextNodeName])=0"/>
				<xsl:otherwise>
					<xsl:call-template name="getContextPosition">
						<xsl:with-param name="nodeList" select="$omigaPath[name()=$contextNodeName]" />
						<xsl:with-param name="contextKey" select="$contextKey" />
						<xsl:with-param name="contextKeyValues" select="$contextKeyValues" />
					</xsl:call-template>				
				</xsl:otherwise>
			</xsl:choose>			
		</xsl:variable>
		
		<!-- Get the omiga node with the same primary key data as the context node. null if no node found -->
		<xsl:variable name="omigaContextNode" select="$omigaPath[name()=$contextNodeName][position()=$omigaContextPosition]" />

		<!-- Get namespace to assign to new element-->
		<xsl:variable name="namespace">
			<xsl:choose>
				<xsl:when test="namespace-uri()!=''"><xsl:value-of select="namespace-uri()" /></xsl:when>
				<xsl:otherwise><xsl:value-of select="'http://OmigaData.Omiga.vertex.co.uk'" /></xsl:otherwise>
			</xsl:choose>
		</xsl:variable>

		<xsl:choose>
			<xsl:when test="count($omigaContextNode)=1">
				<xsl:choose>
					<xsl:when test="@ISCOLLECTION">					
						<!-- 
							TODO:
							May have to iterate through web collection then omiga collection and determine Crud_Ops for all nodes first 
							because the parent node may need to have a CRUD_OP of DELETE OR CREATE if all child collection nodes
							dont exist or are all completely new records. See Ivan
							Will leave CRUD_OP off of parent at the moment
						-->
						<xsl:element name="{local-name()}">
							<!-- Creates and Updates from web data-->
							<!--<xsl:apply-templates select="*">
								<xsl:with-param name="omigaContextNode" select="$omigaContextNode" />
							</xsl:apply-templates>-->
							<!-- Delete's -->
							<xsl:variable name="webContextNode" select="." />
							<xsl:for-each select="$omigaContextNode/*">
								<xsl:apply-templates select="." mode="OmigaCollectionDeletes">
									<xsl:with-param name="webContextNode" select="$webContextNode" />
								</xsl:apply-templates>								
							</xsl:for-each>

							<!-- Process all child elements -->
							<xsl:apply-templates select="*">
								<xsl:with-param name="omigaPath" select="$omigaContextNode/*"/><!-- Append child node to omigaPath-->
							</xsl:apply-templates>

						</xsl:element>
					</xsl:when>
					<xsl:otherwise>
					
						<!-- CRUD_OP = UPDATE -->
						<xsl:variable name="dataHasChanged">
							<xsl:call-template name="hasAttributeDataChanged">
								<xsl:with-param name="omigaAtts" select="$omigaContextNode/@*" />
								<xsl:with-param name="webAtts" select="@*" />
							</xsl:call-template>
						</xsl:variable>
						
						<xsl:if test="$dataHasChanged='true' or count(*)>0">
							<xsl:element name="{local-name()}" namespace="{string($namespace)}">
									<xsl:choose>		
										<!-- TODO: Do I need count(@*) check here? -->
										<xsl:when test="(count(@*)>0 and $dataHasChanged='true')">
											<!-- data has changed copy atts and set UPDATE -->
											<xsl:copy-of select="@*" />									
											<xsl:attribute name="CRUD_OP">
												<xsl:value-of select="'UPDATE'" />
											</xsl:attribute>
										</xsl:when>
										<xsl:otherwise>
											<!-- Data has not changed. Node has child elements so write key data in case referenced by children as keytype="Foreign"-->
											<xsl:copy-of select="@*[name()=$contextKey]" />
										</xsl:otherwise>									
									</xsl:choose>
									
														
								<!-- Iterate child elements -->
								<xsl:apply-templates select="*">
									<xsl:with-param name="omigaPath" select="$omigaContextNode/*"/><!-- Append child node to omigaPath-->
								</xsl:apply-templates>
							</xsl:element>
						</xsl:if>
						<!-- /UPDATE -->
						
					</xsl:otherwise>
				</xsl:choose>
			</xsl:when>
			<xsl:otherwise>
				<!-- CRUD_OP = CREATE -->
				<xsl:element name="{local-name()}" namespace="{string($namespace)}">
					<xsl:if test="count(@*)>0"><!-- If no attributes exist no CRUD_OP=CREATE-->
						<xsl:copy-of select="@*" />
						<xsl:attribute name="CRUD_OP">
							<xsl:value-of select="'CREATE'" />
						</xsl:attribute>
					</xsl:if>
					<xsl:apply-templates select="*">
						<xsl:with-param name="omigaPath" select="$omigaPath/*"/>
					</xsl:apply-templates>				
				</xsl:element>
			</xsl:otherwise>
		</xsl:choose>	

	</xsl:template>
	<!-- /ELEMENTS -->
	
	<!-- COLLECTION Delete's where item is no longer in master XML -->
	<xsl:template match="*" mode="OmigaCollectionDeletes">
		<xsl:param name="webContextNode" />

		<!-- Get string name of the context node. Used to get Key data in contextKey -->
		<xsl:variable name="itemNodeName" select="local-name()" />
		<!-- Get the Primary key attribute names for the context node. Defined in OmigaEntityKeys.xml -->
		<xsl:variable name="itemKey" select="document('OmigaEntityKeys.xml')//ENTITY[@NAME=$itemNodeName]/KEY/@NAME" />
		<!-- Get Nodeset of the actual context nodes key data as defined in $contextKey-->
		<xsl:variable name="omigaKeyValues" select="@*[name()=$itemKey]" />
		<!-- Get Nodeset of omiga nodes under $omigaPath with the same name as the $contextNodeName -->
		<xsl:variable name="webItemNode" select="$webContextNode/*[name()=$itemNodeName][@*=$omigaKeyValues][@*[name()=$itemKey and .=$omigaKeyValues]]" />
		<!-- Get the omiga node with the same primary key data as the context node -->		
	
		<xsl:choose>
			<xsl:when test="count($webItemNode)=0">
				<xsl:element name="{local-name()}">
					<xsl:copy-of select="@*[name()=$itemKey]" />
					<xsl:attribute name="CRUD_OP">
						<xsl:value-of select="'DELETE'" />
					</xsl:attribute>
				</xsl:element>
			</xsl:when>
		</xsl:choose>
	
	</xsl:template>

	
	<xsl:template name="hasAttributeDataChanged">
		<xsl:param name="omigaAtts"  />
		<xsl:param name="webAtts" />
		
		<xsl:variable name="countOfMatches">
			<xsl:call-template name="countMatches">
				<xsl:with-param name="omigaAtts" select="$omigaAtts" />
				<xsl:with-param name="webAtts" select="$webAtts" />
				<xsl:with-param name="matchCount" select="0" />
			</xsl:call-template>
		</xsl:variable>

		<xsl:choose>
			<!-- Were all the web attribute values matched in the omiga nodeset? -->
			<xsl:when test="$countOfMatches!=count($webAtts)">
				<!-- No: data has changed-->
				<xsl:value-of select="'true'" />
			</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select="'false'" />
			</xsl:otherwise>
		</xsl:choose>

	</xsl:template>
	
	
	<xsl:template name="countMatches">
		<xsl:param name="omigaAtts" />
		<xsl:param name="webAtts" />
		<xsl:param name="matchCount" />	
		
		<xsl:choose>
			<xsl:when test="count($webAtts)=0">
				<xsl:value-of select="$matchCount" />
			</xsl:when>
			<xsl:otherwise>
				<xsl:variable name="currentCount">
					<xsl:choose>
						<xsl:when test="$webAtts[position()=1]=$omigaAtts[name()=name($webAtts[position()=1])]">
							<xsl:value-of select="$matchCount +1" />
						</xsl:when>
						<xsl:otherwise>
							<xsl:value-of select="$matchCount" />
						</xsl:otherwise>
					</xsl:choose>
				</xsl:variable>
				
				<xsl:call-template name="countMatches">
					<xsl:with-param name="omigaAtts" select="$omigaAtts" />
					<xsl:with-param name="webAtts" select="$webAtts[position()>1]" />
					<xsl:with-param name="matchCount" select="$currentCount" />
				</xsl:call-template>
				
			</xsl:otherwise>
		</xsl:choose>
		
	</xsl:template>
	
	<!-- All following templates are for locating the associated context node within the supplied nodeList -->
	<!-- Note: I'm sure this code can all be refined in some way -->
	
	<!-- Returns position() of context node in nodeList-->
	<xsl:template name="getContextPosition" match="*" mode="getContextPosition">
		<xsl:param name="nodeList" />
		<xsl:param name="contextKey" />
		<xsl:param name="contextKeyValues" />

		<xsl:for-each select="$nodeList">
			<xsl:variable name="containsAllKeyValues">
				<xsl:call-template name="nodeHasAllKeyValues">
					<xsl:with-param name="node" select="." />
					<xsl:with-param name="contextKey" select="$contextKey" />
					<xsl:with-param name="contextKeyValues" select="$contextKeyValues" />					
				</xsl:call-template>
			</xsl:variable>
			<xsl:if test="$containsAllKeyValues='true'">
				<!-- Returns here -->
				<xsl:value-of select="position()" />
			</xsl:if>
		</xsl:for-each>		
	</xsl:template>	

	<!-- Returns true if node has all the contextKeys with the same contextKeyValues -->
	<xsl:template name="nodeHasAllKeyValues">
		<xsl:param name="node" />
		<xsl:param name="contextKey" />
		<xsl:param name="contextKeyValues" />

		<xsl:variable name="countAtts">
			<xsl:call-template name="countKeys">
				<xsl:with-param name="nodeList" select="$node/@*" />
				<xsl:with-param name="contextKey" select="$contextKey" />
				<xsl:with-param name="contextKeyValues" select="$contextKeyValues" />
				<xsl:with-param name="lastPosition" select="0" />
				<xsl:with-param name="valuesMatchCount" select="0" />
				<xsl:with-param name="contextProcessedCount" select="0" />		
			</xsl:call-template>
		</xsl:variable>

		<xsl:choose>
			<xsl:when test="count($contextKey)=$countAtts">
				<xsl:value-of select="'true'" />
			</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select="'false'" />
			</xsl:otherwise>
		</xsl:choose>		
	</xsl:template>

	<!-- Returns count of attributes that match contextKey ContextKeyValues -->
	<!-- 
		Note: Could speed this up by iterating the contextKey names and matching them in the node.
				 When not found exit immediately.
	-->
	<xsl:template name="countKeys">
		<xsl:param name="nodeList" />
		<xsl:param name="contextKey" />
		<xsl:param name="contextKeyValues" />
		<xsl:param name="lastPosition" />
		<xsl:param name="valuesMatchCount" />
		<xsl:param name="contextProcessedCount" />
		
		<xsl:variable name="currentPosition" select="$lastPosition + 1" />
		<xsl:variable name="currentAtt" select="$nodeList[position()=1]" />

		<xsl:choose>
			<xsl:when test="$contextProcessedCount=count($contextKeyValues)">
				<xsl:value-of select="$valuesMatchCount" />
			</xsl:when>
			<xsl:otherwise>
				<xsl:variable name="vContextProcessedCount">
					<xsl:choose>
						<xsl:when test="name($currentAtt)=$contextKey">
							<xsl:value-of select="$contextProcessedCount + 1" />
						</xsl:when>
						<xsl:otherwise>
							<xsl:value-of select="$contextProcessedCount" />
						</xsl:otherwise>
					</xsl:choose>
				</xsl:variable>			
		
				<xsl:choose>
					<xsl:when test="$currentAtt=$contextKeyValues[name()=name($currentAtt)]">
						<xsl:call-template name="countKeys">
							<xsl:with-param name="nodeList" select="$nodeList[position()>1]" />
							<xsl:with-param name="contextKey" select="$contextKey" />
							<xsl:with-param name="contextKeyValues" select="$contextKeyValues" />
							<xsl:with-param name="lastPosition" select="$currentPosition" />
							<xsl:with-param name="valuesMatchCount" select="$valuesMatchCount + 1"/>		
							<xsl:with-param name="contextProcessedCount" select="$vContextProcessedCount" />		
						</xsl:call-template>
					</xsl:when>
					<xsl:otherwise>
						<xsl:call-template name="countKeys">
							<xsl:with-param name="nodeList" select="$nodeList[position()>1]" />
							<xsl:with-param name="contextKey" select="$contextKey" />
							<xsl:with-param name="contextKeyValues" select="$contextKeyValues" />
							<xsl:with-param name="lastPosition" select="$currentPosition" />
							<xsl:with-param name="valuesMatchCount" select="$valuesMatchCount"/>	
							<xsl:with-param name="contextProcessedCount" select="$vContextProcessedCount" />										
						</xsl:call-template>			
					</xsl:otherwise>
				</xsl:choose>			
			</xsl:otherwise>
		</xsl:choose>

	</xsl:template>		
	
	
			
</xsl:stylesheet>
