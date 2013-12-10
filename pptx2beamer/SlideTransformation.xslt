<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:msxsl="urn:schemas-microsoft-com:xslt" exclude-result-prefixes="msxsl"
                xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                xmlns:x="urn:pptx2beamer-addition"
>
    <xsl:output method="text" indent="yes"/>
    <xsl:template match="@* | node()">
\frame{<xsl:if test="//p:ph[@type='title']" >
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
      <xsl:for-each select="//p:pic">
  % Image contents
  \begin{figure}
  \includegraphics{img/<xsl:value-of select="p:blipFill/a:blip/@src" />}
  \end{figure}
      </xsl:for-each>    
}
    </xsl:template>
</xsl:stylesheet>