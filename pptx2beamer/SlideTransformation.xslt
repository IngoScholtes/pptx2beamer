<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:msxsl="urn:schemas-microsoft-com:xslt" exclude-result-prefixes="msxsl"
                xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
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
    \frametitle{
    <xsl:for-each select="//p:ph[@type='title']/../../../p:txBody/a:p/a:r">
      <xsl:variable name="t" select="a:t" />      
        <xsl:value-of select="a:t" />
      <xsl:if test="substring($t, string-length($t),1)!=' '">
        <xsl:text> </xsl:text>
      </xsl:if>
      
    </xsl:for-each>
    }
  </xsl:if>
  <xsl:if test="//p:cNvPr[starts-with(@name, 'Inhaltsplatzhalter')]">
  %  Text contents
  \begin{itemize}
  <xsl:for-each select="//p:cNvPr[starts-with(@name, 'Inhaltsplatzhalter')]/../../p:txBody/a:p/a:r">
    \item <xsl:value-of select="a:t" /> \\
  </xsl:for-each>  
  \end{itemize}
  </xsl:if>

  <xsl:if test="//p:pic">
  % Image contents
  </xsl:if>
}
    </xsl:template>
</xsl:stylesheet>