<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:msxsl="urn:schemas-microsoft-com:xslt" exclude-result-prefixes="msxsl"
                 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
>
    <xsl:output method="text" indent="yes"/>

    <xsl:template match="@* | node()">
      <xsl:if test="//p:cNvPr[starts-with(@name, 'Notizenplatzhalter')]">
  % slide notes
  \note{
        <xsl:for-each select="//p:cNvPr[starts-with(@name, 'Notizenplatzhalter')]/../../p:txBody/a:p/a:r">
          <xsl:variable name="t" select="a:t" />
          <xsl:value-of select="a:t" />
          <xsl:if test="substring($t, string-length($t),1)!=' '">
            <xsl:text> </xsl:text>
          </xsl:if>
        </xsl:for-each>
  }
      </xsl:if>
    </xsl:template>
</xsl:stylesheet>
