<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:msxsl="urn:schemas-microsoft-com:xslt" exclude-result-prefixes="msxsl"
                xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                xmlns:x="urn:pptx2beamer-addition"
>

  <xsl:output method="text" indent="no"/>
  <!-- A template which replace any occurrence of a substring by another substring -->
  <xsl:template name="string-replace-all">
    <xsl:param name="text" />
    <xsl:param name="replace" />
    <xsl:param name="by" />
    <xsl:choose>
      <xsl:when test="contains($text, $replace)">
        <xsl:value-of select="substring-before($text,$replace)" />
        <xsl:value-of select="$by" />
        <xsl:call-template name="string-replace-all">
          <xsl:with-param name="text" select="substring-after($text,$replace)" />
          <xsl:with-param name="replace" select="$replace" />
          <xsl:with-param name="by" select="$by" />
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$text" />
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <!-- A template which escapes all special LaTeX characters-->
  <xsl:template name="escape-latex">
    <xsl:param name="text" />
    <xsl:variable name="escape1" >
      <xsl:call-template name="string-replace-all">
        <xsl:with-param name="text" select="$text" />
        <xsl:with-param name="replace" select="'%'" />
        <xsl:with-param name="by" select="'\%'" />
      </xsl:call-template>      
    </xsl:variable>
    <xsl:variable name="escape2" >
      <xsl:call-template name="string-replace-all">
        <xsl:with-param name="text" select="$escape1" />
        <xsl:with-param name="replace" select="'$'" />
        <xsl:with-param name="by" select="'\$'" />
      </xsl:call-template>
    </xsl:variable>
    <xsl:value-of select="$escape2"/>
  </xsl:template>
  <!-- A template which transforms elements on a pptx slide to LaTeX beamer code  -->
  <xsl:template match="@* | node()">
    <!-- Generate frame and extract frametitle -->    
  \frame{<xsl:if test="//p:ph[@type='title']" >
      \frametitle{<xsl:for-each select="//p:ph[@type='title']/../../../p:txBody/a:p/a:r">        
        <xsl:variable name="t" select="a:t" />      
          <xsl:value-of select="$t" /> 
        <xsl:if test="substring($t, string-length($t),1)!=' '">
            <xsl:text> </xsl:text>
        </xsl:if>        
      </xsl:for-each>}
  </xsl:if>    
  <!-- Transform text placeholder elements -->
  <xsl:for-each select="//p:sp[p:nvSpPr/p:cNvPr[starts-with(@name, 'Inhaltsplatzhalter')]]">
  <!-- Check for multi-column layout -->    
      <xsl:if test="p:nvSpPr/p:nvPr/p:ph[@sz='half']">
        <!--  Start of multi-column layout if no preceding content placeholder with sz=half-->
        <xsl:if test="not(preceding-sibling::p:sp[p:nvSpPr/p:nvPr/p:ph[@sz='half']])">
    \begin{columns}
        </xsl:if>
        \column{0.45\textwidth}
      </xsl:if>
<!-- Generate itemized list -->        
        \begin{itemize}
      <xsl:for-each select="p:txBody/a:p">        
        \item  <xsl:for-each select="a:r">
          <xsl:if test ="a:rPr/a:solidFill"> { \color[HTML]{<xsl:value-of select="a:rPr/a:solidFill/a:srgbClr/@val"/>} </xsl:if>
            <xsl:variable name="newtext" >
              <xsl:call-template name="escape-latex">
                <xsl:with-param name="text" select="a:t"/> 
              </xsl:call-template>
            </xsl:variable>
          <xsl:value-of select="$newtext"/>
          <xsl:value-of select="' '"/>
          <xsl:if test ="a:rPr/a:solidFill"> } </xsl:if>
          </xsl:for-each> \\
  </xsl:for-each>
        \end{itemize}
  <!-- End of multi-column layout if no following content placeholder with sz=half-->        
      <xsl:if test="p:nvSpPr/p:nvPr/p:ph[@sz='half'] and not(following-sibling::p:sp[p:nvSpPr/p:nvPr/p:ph[@sz='half'] and p:nvSpPr/p:cNvPr[starts-with(@name, 'Inhaltsplatzhalter')] ])">      
    \end{columns}
      </xsl:if>
  </xsl:for-each>  
  <!-- Transform pictures -->
  <xsl:for-each select="//p:pic">
          <xsl:variable name="w" select="p:spPr/a:xfrm/a:ext/@cx div 828000"/>
          <xsl:variable name="h" select="p:spPr/a:xfrm/a:ext/@cy div 828000"/>
          <xsl:variable name="x" select="p:spPr/a:xfrm/a:off/@x  div 828000"/>
          <xsl:variable name="y" select="p:spPr/a:xfrm/a:off/@y  div 828000"/>
      \begin{textblock*}{<xsl:value-of select="$w"/>cm}(<xsl:value-of select="$x"/>cm,<xsl:value-of select="$y"/>cm)
        \begin{figure}  
          \includegraphics[width=<xsl:value-of select="$w"/>cm,height=<xsl:value-of select="$h"/>cm]{img/<xsl:value-of select="p:blipFill/a:blip/@src" />}
        \end{figure}
      \end{textblock*}
    </xsl:for-each>
    <!-- Transform randomly positioned text boxes -->
    <xsl:for-each select="//p:sp/p:nvSpPr/p:cNvSpPr[@txBox='1']">
      <xsl:variable name="w" select="../../p:spPr/a:xfrm/a:ext/@cx div 828000"/>
      <xsl:variable name="h" select="../../p:spPr/a:xfrm/a:ext/@cy div 828000"/>
      <xsl:variable name="x" select="../../p:spPr/a:xfrm/a:off/@x  div 828000"/>
      <xsl:variable name="y" select="../../p:spPr/a:xfrm/a:off/@y  div 828000"/>
      \begin{textblock*}{<xsl:value-of select="$w"/>cm}(<xsl:value-of select="$x"/>cm,<xsl:value-of select="$y"/>cm)      
      <xsl:for-each select="../../p:txBody/a:p">
        <xsl:for-each select="a:r">
          <xsl:variable name="newtext" >
            <xsl:call-template name="escape-latex">
              <xsl:with-param name="text" select="a:t"/>
            </xsl:call-template>
          </xsl:variable>
          <xsl:value-of select="$newtext"/>
          <xsl:value-of select="' '"/>
        </xsl:for-each>             
       </xsl:for-each>
      \end{textblock*}
    </xsl:for-each>}
  </xsl:template>
</xsl:stylesheet>