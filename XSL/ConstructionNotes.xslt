<?xml version="1.0" encoding="UTF-8" ?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
  <xsl:key name="CU-by-Wl-and-Function" match="//WORKLOCATION//CU" use="concat(ancestor::node()[local-name()='WORKLOCATION']/ID,' ',WORK_FUNCTION)"/>
  <xsl:key name="CU-by-Dn-and-Function" match="//DESIGN/GISUNIT//CU|//DESIGN/CU" use="concat(ancestor::node()[local-name()='DESIGN']/ID,' ',WORK_FUNCTION)"/>
  <xsl:key name="WlCU-by-WF-CODE-NOTE" match="//WORKLOCATION//CU" use="concat(ancestor::node()[local-name()='WORKLOCATION']/ID,' ',WORK_FUNCTION,' ',WMS_CODE)"/>
  <xsl:key name="DnCU-by-WF-CODE-NOTE" match="//DESIGN/GISUNIT//CU|//DESIGN/CU" use="concat(ancestor::node()[local-name()='DESIGN']/ID,' ',WORK_FUNCTION,' ',WMS_CODE)"/>
  <xsl:template match="/">
    <xsl:for-each select="//WORKLOCATION">
      <xsl:sort select="ID"/>
      <xsl:apply-templates select="."/>
      <xsl:for-each select=".//CU[ancestor::node()[local-name()='WORKLOCATION']]">
        <xsl:sort select="WORK_FUNCTION"/>
        <xsl:apply-templates select="." />
      </xsl:for-each>
    </xsl:for-each>
    <!--Explicitly ignore unassigned CUs-->
  </xsl:template>
  <xsl:template match="WORKLOCATION" name="Wl">
    <BOL>
      WL <xsl:value-of select="ID"/><xsl:text> </xsl:text><xsl:value-of select="DESCRIPTION"/><xsl:text>: </xsl:text><xsl:value-of select="EDM/EDMPROP[@Name='INSTALL_NOTES']"/>
    </BOL>
    <xsl:text> 
</xsl:text>
  </xsl:template>
  <xsl:template match="CU">
        <xsl:for-each select="self::node()[count(.|key('WlCU-by-WF-CODE-NOTE',concat(ancestor::node()[local-name()='WORKLOCATION']/ID,' ',WORK_FUNCTION,' ',WMS_CODE))[1])=1]">
          <!--Test to Display work function-->
          <xsl:if test="count(.|key('CU-by-Wl-and-Function',concat(ancestor::node()[local-name()='WORKLOCATION']/ID,' ',WORK_FUNCTION))[1])=1">
            <xsl:apply-templates select="WORK_FUNCTION"/>
          </xsl:if>
          <!--Display quantity or length-->
          <xsl:choose>
            <xsl:when test="TABLENAME and ../LENGTH">(<xsl:value-of select="round(sum(key('WlCU-by-WF-CODE-NOTE',concat(ancestor::node()[local-name()='WORKLOCATION']/ID,' ',WORK_FUNCTION,' ',WMS_CODE))/../LENGTH))"/>)</xsl:when>
            <xsl:when test="QUANTITY">(<xsl:value-of select="sum(key('WlCU-by-WF-CODE-NOTE',concat(ancestor::node()[local-name()='WORKLOCATION']/ID,' ',WORK_FUNCTION,' ',WMS_CODE))/QUANTITY )"/>)</xsl:when>
            <xsl:otherwise>(1)</xsl:otherwise>
          </xsl:choose>
          <!--Display CUName-->
          <xsl:value-of select="CUNAME"/>
          <!--<xsl:text>: </xsl:text>
          <xsl:value-of select="EDM/EDMPROP[@Name='INSTALL_NOTES']"/>-->
          <xsl:text>
</xsl:text>
        </xsl:for-each>
  </xsl:template>
  <xsl:template name="WF" match="WORK_FUNCTION">
    <xsl:choose>
      <xsl:when test="text()='1'">Install</xsl:when>
      <xsl:when test="text()='2'">Remove</xsl:when>
      <xsl:when test="text()='4'">Salvage</xsl:when>
      <xsl:when test="text()='8'">Abandon</xsl:when>
      <xsl:otherwise>
        <xsl:text>Unknown</xsl:text>
      </xsl:otherwise>
    </xsl:choose>
    <xsl:text>: </xsl:text>
    <xsl:text>
</xsl:text>
  </xsl:template>
</xsl:stylesheet>