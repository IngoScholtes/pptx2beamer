<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:msxsl="urn:schemas-microsoft-com:xslt" exclude-result-prefixes="msxsl"
                xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
>
    <xsl:output method="text" indent="yes"/>

    <xsl:template match="@* | node()">
\frame{
  <xsl:if test="//p:ph[@type='ctrTitle']">
    \title{<xsl:value-of select="//p:ph[@type='ctrTitle']/../../../p:txBody"/>}
    \author{Ingo Scholtes}
    \date{today}
    \institute{Chair of Systems Design}
    \maketitle
  </xsl:if>
  <xsl:if test="//p:ph[@type='title']" >
    \frametitle{<xsl:value-of select="//p:ph[@type='title']/../../../p:txBody" /> }
  </xsl:if>
  <xsl:if test="//p:cNvPr[starts-with(@name, 'Inhaltsplatzhalter')]">
  %  Text contents
  <xsl:value-of select="//p:cNvPr[starts-with(@name, 'Inhaltsplatzhalter')]/../../p:txBody" />
  </xsl:if>

  <xsl:if test="//p:pic">
  % Image contents
  </xsl:if>
}
    </xsl:template>
</xsl:stylesheet>