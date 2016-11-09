<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet xmlns:xs="http://www.w3.org/2001/XMLSchema"
                xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                xmlns:prop="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                xmlns:dc="http://purl.org/dc/elements/1.1/"
                xmlns:dcterms="http://purl.org/dc/terms/"
                xmlns:dcmitype="http://purl.org/dc/dcmitype/"
                xmlns:iso="http://www.iso.org/ns/1.0"
                xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
                xmlns:mml="http://www.w3.org/1998/Math/MathML"
                xmlns:mo="http://schemas.microsoft.com/office/mac/office/2008/main"
                xmlns:mv="urn:schemas-microsoft-com:mac:vml"
                xmlns:o="urn:schemas-microsoft-com:office:office"
                xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships"
                xmlns:tbx="http://www.lisa.org/TBX-Specification.33.0.html"
                xmlns:tei="http://www.tei-c.org/ns/1.0"
                xmlns:teidocx="http://www.tei-c.org/ns/teidocx/1.0"
                xmlns:v="urn:schemas-microsoft-com:vml"
                xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006"
                xmlns:w10="urn:schemas-microsoft-com:office:word"
                xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
                xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                
                xmlns:foo="http://foo"
                
                xmlns="http://www.tei-c.org/ns/1.0"
                version="2.0"
                exclude-result-prefixes="a cp dc dcterms dcmitype prop     iso m mml mo mv o pic r rel     tbx tei teidocx v xs ve w10 w wne wp">
    
    <doc xmlns="http://www.oxygenxml.com/ns/doc/xsl" scope="stylesheet" type="stylesheet">
      <desc>
         <p> TEI stylesheet for converting Word docx files to TEI </p>
         
         <p> Note: This lists.xsl now bears little resemblance to the original TEI file of that name.</p>
         
         <p>The TEI stylesheet is dual-licensed:

1. Distributed under a Creative Commons Attribution-ShareAlike 3.0
Unported License http://creativecommons.org/licenses/by-sa/3.0/ 

2. http://www.opensource.org/licenses/BSD-2-Clause
		


Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are
met:

* Redistributions of source code must retain the above copyright
notice, this list of conditions and the following disclaimer.

* Redistributions in binary form must reproduce the above copyright
notice, this list of conditions and the following disclaimer in the
documentation and/or other materials provided with the distribution.

This software is provided by the copyright holders and contributors
"as is" and any express or implied warranties, including, but not
limited to, the implied warranties of merchantability and fitness for
a particular purpose are disclaimed. In no event shall the copyright
holder or contributors be liable for any direct, indirect, incidental,
special, exemplary, or consequential damages (including, but not
limited to, procurement of substitute goods or services; loss of use,
data, or profits; or business interruption) however caused and on any
theory of liability, whether in contract, strict liability, or tort
(including negligence or otherwise) arising in any way out of the use
of this software, even if advised of the possibility of such damage.
</p>
         <p>Author: See AUTHORS</p>
         
         <p>Copyright: 2013, TEI Consortium</p>
      </desc>
   </doc>
   
   <!--  NOT WORKING?
   
   Check:
   
   1.  your input docx has a heading
   2.  you are writing item content
   
    -->
    
    <xsl:template match="w:p" mode="makeItem">
    
    	<foo:p>
	        <xsl:variable name="style">
	            <xsl:value-of select="w:pPr/w:pStyle/@w:val"/>
	        </xsl:variable>
	        <xsl:attribute name="rend">
	        	<xsl:value-of select="$style"/>
	        </xsl:attribute>
	                    
	         <xsl:variable name="numbering-def">
	             <xsl:choose>
	                 <xsl:when test="w:pPr/w:numPr/w:numId/@w:val">
	                     <xsl:value-of select="w:pPr/w:numPr/w:numId/@w:val"/>
	                 </xsl:when>
	                 <xsl:otherwise>
	                     <xsl:value-of select="document($styleDoc)//w:style[@w:styleId=$style]/w:pPr/w:numPr/w:numId/@w:val"/>
	                 </xsl:otherwise>
	             </xsl:choose>
	         </xsl:variable>
	         
			<xsl:attribute name="numId"> 
		           <xsl:value-of select="$numbering-def"/>
			</xsl:attribute>
	         
	         
	         <xsl:variable name="numbering-level">
	             <xsl:choose>
	                 <xsl:when test="w:pPr/w:numPr/w:ilvl/@w:val">
	                     <xsl:value-of select="w:pPr/w:numPr/w:ilvl/@w:val"/>
	                 </xsl:when>
	                 <xsl:when test="w:pPr/w:numPr">0</xsl:when>
	                 <xsl:otherwise>
			               <xsl:choose>
			               	<xsl:when test="document($styleDoc)//w:style[@w:styleId=$style]/w:pPr/w:numPr/w:ilvl">
			                       <xsl:value-of select="document($styleDoc)//w:style[@w:styleId=$style]/w:pPr/w:numPr/w:ilvl/@w:val"/>
			               	</xsl:when>
			               	<xsl:when test="document($styleDoc)//w:style[@w:styleId=$style]/w:pPr/w:numPr">0</xsl:when>
			                   <xsl:otherwise/>
			               </xsl:choose>
	                 </xsl:otherwise>
	             </xsl:choose>
	         </xsl:variable>
		      
		     <!--             
	        <xsl:attribute name="ilvl">
	            <xsl:value-of select="$numbering-level"/>
			</xsl:attribute>
	         -->
		                 
	         <xsl:variable name="abstract-def"
	                  select="document($numberFile)//w:num[@w:numId=$numbering-def]/w:abstractNumId/@w:val"/>
	         
	         <!--  access <w:ind w:left="720" w:hanging="360"/> 
	         
	         	   we use @w:left to infer list level
	          -->         
	         <xsl:variable name="indleft">
	             <xsl:value-of select="document($numberFile)//w:abstractNum[@w:abstractNumId=$abstract-def]/w:lvl[@w:ilvl=$numbering-level]/w:pPr//w:ind/@w:left"/>
	         </xsl:variable>
	                  
		     <xsl:attribute name="indleft"> <!--  just for debugging -->
		           <xsl:value-of select="$indleft"/>
			</xsl:attribute>
			
			<xsl:variable name="thisLevel"><xsl:value-of select="number(foo:level(
																		number($indleft)  ))"/></xsl:variable>

			<xsl:attribute name="ilvl"> 
		           <xsl:value-of select="$thisLevel"/>
			</xsl:attribute>
			
	                  
	         <xsl:variable name="numfmt">
	             <xsl:value-of select="document($numberFile)//w:abstractNum[@w:abstractNumId=$abstract-def]/w:lvl[@w:ilvl=$numbering-level]/w:numFmt/@w:val"/>
	         </xsl:variable>
		                 
		       <xsl:attribute name="numFmt">
		           <xsl:value-of select="$numfmt"/>
			</xsl:attribute>
		      
		     <!--             
	         <xsl:variable name="type">
	             <xsl:choose>
			        <xsl:when test="$style='dl'">gloss</xsl:when>
	                      <xsl:when test="string-length($numfmt)=0">unordered</xsl:when>
	                      <xsl:when test="$numfmt='bullet'">unordered</xsl:when>
	                      <xsl:otherwise>ordered</xsl:otherwise>
	             </xsl:choose>
	        </xsl:variable>
	        <xsl:attribute name="type"><xsl:value-of select="$type"/></xsl:attribute>
	         -->	
	         
	         <xsl:copy-of select="@*|node()"/>          
    	</foo:p>
	         
	</xsl:template>	        
	        


    
    <!-- 
        This template handles lists and takes care of nested lists.
    -->
    <xsl:template name="listSection">
    	<xsl:param name="items"/>
    	
         <xsl:variable name="list">
		    	<xsl:for-each select="$items">
			         	<xsl:apply-templates select="." mode="makeItem"/>
		    	</xsl:for-each>    	
         </xsl:variable>
         
         <xsl:call-template name="nest">
         	<xsl:with-param name="pSequence" select="$list"/>
         </xsl:call-template>

    </xsl:template>
    


	<!--  The idea of this function is to infer the level of the list from its indentation:

		    <w:lvl w:ilvl="0" 
		        <w:ind w:left="720" 
		
		    <w:lvl w:ilvl="1" 
		        <w:ind w:left="1440" 
		
		    <w:lvl w:ilvl="2" 
		        <w:ind w:left="2160" 
		
		    <w:lvl w:ilvl="3" 
		        <w:ind w:left="2880" 
		
		    <w:lvl w:ilvl="4" 
		        <w:ind w:left="3600" 
		
		    <w:lvl w:ilvl="5" 
		        <w:ind w:left="4320" 
		
		    <w:lvl w:ilvl="6" 
		        <w:ind w:left="5040" 
		
		    <w:lvl w:ilvl="7" 
		        <w:ind w:left="5760"
		
		    <w:lvl w:ilvl="8" 
		        <w:ind w:left="6480" 


        -->
    <xsl:function name="foo:level" as="xs:integer">
        <xsl:param name="indleft"/> 
        
           
            <xsl:choose>
                <xsl:when test="$indleft &lt;= 720">0</xsl:when>
                <xsl:when test="$indleft &lt;= 1440">1</xsl:when>
                <xsl:when test="$indleft &lt;= 2160">2</xsl:when>
                <xsl:when test="$indleft &lt;= 2880">3</xsl:when>
                <xsl:when test="$indleft &lt;= 3600">4</xsl:when>
                <xsl:when test="$indleft &lt;= 4320">5</xsl:when>
                <xsl:when test="$indleft &lt;= 5040">6</xsl:when>
                <xsl:when test="$indleft &lt;= 5760">7</xsl:when>
                <xsl:otherwise>8</xsl:otherwise>
            </xsl:choose>
            
    </xsl:function>    
     


 <xsl:param name="pStartLevel" select="0"/>
 <!--  Note: this assumes first entry is ilvl 0, which the Word interface enforces, so should be ok -->


 <!--  approach below devised by http://stackoverflow.com/users/36305/dimitre-novatchev -->

 <xsl:key name="kRByLevelAndParent" match="foo:p"
  use="generate-id(preceding-sibling::foo:p
                            [not(@ilvl >= current()/@ilvl)][1])"/>

 <xsl:template name="nest">
   <xsl:param name="pSequence"/>
 
      <xsl:apply-templates select="key('kRByLevelAndParent', '', $pSequence)[1]" mode="start">
        <xsl:with-param name="pParentLevel" select="$pStartLevel"/>
        <xsl:with-param name="pSiblings" select="key('kRByLevelAndParent', '', $pSequence)"/>
      </xsl:apply-templates>
 </xsl:template>

 <xsl:template match="foo:p" mode="start">
   <xsl:param name="pParentLevel"/>
   <xsl:param name="pSiblings"/>

   
   <list>
    <xsl:apply-templates select="$pSiblings">
      <xsl:with-param name="pParentLevel" select="$pParentLevel"/>
    </xsl:apply-templates>
  </list>
 </xsl:template>

 <xsl:template match="foo:p">
   <xsl:param name="pParentLevel"/>

   <xsl:choose>
       <xsl:when test="@ilvl - $pParentLevel > 1">
       <item>
         <!-- <p ilvl="{$pParentLevel +1}"/>  -->
         <xsl:comment>interpolated level</xsl:comment>
         <list>
           <xsl:apply-templates select=".">
             <xsl:with-param name="pParentLevel" select="$pParentLevel +1"/>
           </xsl:apply-templates>
         </list>
       </item>
       </xsl:when>
       <xsl:otherwise>
           <item>
			 <!--  copy our attributes -->
			         
			<!--  uncomment for debug 
		     <xsl:attribute name="indleft"> 
		           <xsl:value-of select="@indleft"/>
			</xsl:attribute>	
			-->
			
			<!--  inferred/intended level: might be different from actual w:ilvl -->		
			<xsl:attribute name="ilvl"> 
		           <xsl:value-of select="@ilvl"/>
			</xsl:attribute>
			
		    <xsl:attribute name="rend">
		           <xsl:value-of select="@rend"/>
			</xsl:attribute>  
			
		    <xsl:attribute name="numFmt">
		           <xsl:value-of select="@numFmt"/>
			</xsl:attribute>  

			 <!--  where @rend and @numFmt contradict, numbering has probably been restarted.
			 
			 	   @rend is from the style, which will reflect the original list
			 	   
			 	   @numFmt is the format in the "buggy" new list Word created automatically,
			 	   and although it reflects what the user sees in Word, it is not what we intended.
			 	   
			 	   So you probably want to use the formatting hint in @rend 
			 
			  -->
			 
			 <!--  a newly encountered numId means numbering has been restarted -->
			<xsl:attribute name="numId"> 
		           <xsl:value-of select="@numId"/>
			</xsl:attribute>
			 
			 <!--  handle contents of this item -->        
             <xsl:apply-templates/>
             
			 <!--  continue -->        
             <xsl:apply-templates mode="start"
                  select="key('kRByLevelAndParent',generate-id())[1]">
               <xsl:with-param name="pParentLevel" select="@ilvl"/>
               <xsl:with-param name="pSiblings" 
                   select="key('kRByLevelAndParent',generate-id())"/>
             </xsl:apply-templates>
           </item>
       </xsl:otherwise>
   </xsl:choose>
 </xsl:template>    
</xsl:stylesheet>
