<?xml version='1.0'?>
<xsl:stylesheet version="2.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">

    <!--
        Excel 2012 generic files generator

        This file generates the 'generic' files, that are needed in order to generate
        valid excel file.

        It is included and should not be used directly.

        @author Pavel Ptacek
        @copyright Pavel Ptacek (c) 2012-2013
    -->

    <!-- generate theme file -->
    <xsl:template name="generate_themes" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <xsl:param name="content" />

        <xsl:for-each select="$content/a:theme">
            <xsl:result-document href="xl/theme/theme{position()}.xml">
                <xsl:copy-of select="." />
            </xsl:result-document>
        </xsl:for-each>
    </xsl:template>

    <!-- generate styles file -->
    <xsl:template name="generate_styles">
        <xsl:param name="content" />

        <xsl:result-document href="xl/styles.xml">
            <xsl:copy-of select="$content" />
        </xsl:result-document>
    </xsl:template>

    <!-- generate binary file -->
    <xsl:template name="generate_binary">
        <xsl:param name="name"/>
        <xsl:result-document href="{$name}" media-type="text/plain" omit-xml-declaration="yes">
            <xsl:fallback/>
        </xsl:result-document>
    </xsl:template>

</xsl:stylesheet>