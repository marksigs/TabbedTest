<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
   
   <xsl:output method="html" indent="yes" />

   <xsl:template match="menu">
         <ul class="sidemenu">
           <xsl:apply-templates />
         </ul>
   </xsl:template>
   
   <xsl:template match="menuitem">
     <xsl:choose>
       <xsl:when test="*">
         <xsl:element name="li">
           <xsl:if test="@class">
             <xsl:attribute name="class"><xsl:value-of select="@class" /></xsl:attribute>
           </xsl:if>
           <xsl:value-of select="@name" />
           <ul>
             <xsl:apply-templates />
           </ul>
         </xsl:element>
       </xsl:when>
       <xsl:otherwise>
         <xsl:element name="li">
           <xsl:if test="@class">
             <xsl:attribute name="class"><xsl:value-of select="@class" /></xsl:attribute>
           </xsl:if>
           <xsl:value-of select="@name" />
         </xsl:element>
       </xsl:otherwise>
     </xsl:choose>
   </xsl:template>
   
</xsl:stylesheet>