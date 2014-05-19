<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:xs="http://www.w3.org/2001/XMLSchema"
    xmlns:xd="http://www.oxygenxml.com/ns/doc/xsl"
    xmlns:s="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    exclude-result-prefixes="#all"
    version="2.0">
    <xd:doc scope="stylesheet">
        <xd:desc>
            <xd:p><xd:b>This file is part of the xsl-excel-engine github project</xd:b></xd:p>
            <xd:p>This XSLT performs the same function as sharedStringsConverter.php, just implemented in XSLT so that you don't have to have PHP. For information about what this does read the documentation. How to use:</xd:p>
            <xd:ul>
                <xd:li>extract your .xslx file as zip into folder</xd:li>
                <xd:li>provide parameter sharedStringsXml with the path to sharedStrings.xml (by default assumed to be in the same folder)</xd:li>
                <xd:li>pick your sheetX.xml file and provide it as input</xd:li>
                <xd:li>open the output file in your favorite XML editor and use it to create your worksheet template in XSLT</xd:li>
            </xd:ul>
        </xd:desc>
    </xd:doc>
    
    <xsl:param name="sharedStringsXml" select="'sharedStrings.xml'"/>
    
    <xsl:variable name="sharedStrings" select="doc($sharedStringsXml)"/>
    
    <xsl:template match="@* | node()">
        <xsl:copy>
            <xsl:apply-templates select="@* | node()"/>
        </xsl:copy>
    </xsl:template>
    
    <xsl:template match="s:c[@t eq 's']">
        <xsl:element name="c" namespace="{namespace-uri()}">
            <xsl:apply-templates select="@*[not(local-name() eq 't')]"/>
            <xsl:attribute name="t" select="'inlineStr'"/>
            <xsl:element name="is" namespace="{namespace-uri()}">
                <xsl:element name="t" namespace="{namespace-uri()}">
                    <xsl:variable name="v" select="number(s:v)"/>
                    <xsl:value-of select="$sharedStrings/s:sst/s:si[count(preceding-sibling::s:si) eq $v]/s:t"/>
                </xsl:element>
            </xsl:element>
        </xsl:element>
    </xsl:template>
    
</xsl:stylesheet>