<?xml version='1.0'?>
<xsl:stylesheet version="2.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">

    <!--
            The main entry point of Excel 2012 generator

            You should supply all parameters which you want to alter, or are required.
            The generation will take place automatically.

            Copy-paste parameters:
            - themes (xmlNodes from xl/themes/theme1.xml)
            - styles (xmlNodes, content of xl/styles.xml)

            Other parameters:
            - author (string, optional, default "YOURCOMPANY Team")

            - sheetNames (xml nodes)
                <name>Sheet</name>
                <name>Sheet</name>

            - sheetContents (xmlNodes - - contents of the sheetX.xml files!)
                <worksheet>...</worksheet>
                <worksheet>...</worksheet>

            - vbas (xmlNodes - used for macro-enabled workbooks)
                <vba>filename.bin</vba>
                <vba>filename.bin</vba>

            - printerSettings (xmlNodes)
                <setting for="1">sheet1_settings.bin</setting>
                <setting for="2">sheet2_settings.bin</setting>
                -> printerSettings is always rId2 within the worksheet

            - images (xmlNodes)
                - the usedIn parameter is required to list sheet X numbers (nr-ed from 1)
                - the xdr:from, xdr:to, xdr:spPr is copy-paste from xl/drawings/drawing1.xml

                example code:
                <image>
                    <path>blank_image</path>
                    <ext>jpg</ext>
                    <mime>image/jpeg</mime>
                    <usedIn>
                        <sheet nr="1" />
                    </usedIn>
                    <xdr:from>
                      <xdr:col>1</xdr:col>
                      <xdr:colOff>10783</xdr:colOff>
                      <xdr:row>2</xdr:row>
                      <xdr:rowOff>10784</xdr:rowOff>
                    </xdr:from>
                    <xdr:to>
                      <xdr:col>4</xdr:col>
                      <xdr:colOff>596660</xdr:colOff>
                      <xdr:row>10</xdr:row>
                      <xdr:rowOff>190500</xdr:rowOff>
                    </xdr:to>
                    <xdr:spPr>
                      <a:xfrm>
                        <a:off x="621821" y="2595114"/>
                        <a:ext cx="2418990" cy="1703716"/>
                      </a:xfrm>
                      <a:prstGeom prst="rect">
                        <a:avLst/>
                      </a:prstGeom>
                    </xdr:spPr>
                </image>

            @author Pavel Ptacek
            @copyright Pavel Ptacek (c) 2012-2013
    -->

    <!-- required includes -->
    <xsl:include href="2012.defaultdata.xsl" />
    <xsl:include href="2012.generic.xsl" />
    <xsl:include href="2012.manifest.xsl" />
    <xsl:include href="2012.properties.xsl" />
    <xsl:include href="2012.media.xsl" />
    <xsl:include href="2012.worksheet.xsl" />

    <!-- The black magic template -->
    <xsl:template name="generate_excel">

        <!-- Properties of the excel file -->
        <xsl:param name="author">xsl-excel-engine</xsl:param>
        <xsl:param name="sheetNames"><name>Sheet1</name></xsl:param>

        <!-- The drawings you are using in worksheets -->
        <xsl:param name="images" select="()"/>

        <!-- default parameters -->
        <xsl:param name="themes"><xsl:copy-of select="$__blank_theme" /></xsl:param>
        <xsl:param name="styles"><xsl:copy-of select="$__blank_styles" /></xsl:param>
        <xsl:param name="sheetContents"><xsl:copy-of select="$__blank_sheet" /></xsl:param>
        <xsl:param name="vbas" />
        <xsl:param name="printerSettings" />

        <!-- DO NOT TOUCH BELOW THIS LINE -->

        <!-- xl/theme/themeX.xml -->
        <xsl:call-template name="generate_themes">
            <xsl:with-param name="content"><xsl:copy-of select="$themes" /></xsl:with-param>
        </xsl:call-template>

        <!-- xl/styles.xml -->
        <xsl:call-template name="generate_styles">
            <xsl:with-param name="content"><xsl:copy-of select="$styles" /></xsl:with-param>
        </xsl:call-template>

        <!-- Generate manifests -->
        <xsl:variable name="media"><xsl:if test="$images"><xsl:for-each select="$images/image">
            <image>
                <ext><xsl:value-of select="./ext" /></ext>
                <mime><xsl:value-of select="./mime" /></mime>
            </image>
        </xsl:for-each></xsl:if></xsl:variable>
        <xsl:call-template name="generate_manifest">
            <xsl:with-param name="sheets"><xsl:value-of select="count($sheetNames/name)" /></xsl:with-param>
            <xsl:with-param name="drawings"><xsl:value-of select="count($images/image/usedIn/sheet)" /></xsl:with-param>
            <xsl:with-param name="media"><xsl:copy-of select="$media" /></xsl:with-param>
            <xsl:with-param name="vbas"><xsl:copy-of select="$vbas" /></xsl:with-param>
        </xsl:call-template>

        <!-- Generate properties -->
        <xsl:call-template name="generate_properties">
            <xsl:with-param name="author"><xsl:value-of select="$author" /></xsl:with-param>
            <xsl:with-param name="sheetNames"><xsl:copy-of select="$sheetNames" /></xsl:with-param>
        </xsl:call-template>

        <!-- Generate images where applicable -->
        <xsl:call-template name="generate_media">
            <xsl:with-param name="images"><xsl:copy-of select="$images" /></xsl:with-param>
            <xsl:with-param name="printerSettings"><xsl:copy-of select="$printerSettings" /></xsl:with-param>
        </xsl:call-template>

        <!-- Generate the sheet -->
        <xsl:call-template name="generate_sheets">
            <xsl:with-param name="sheetNames"><xsl:copy-of select="$sheetNames" /></xsl:with-param>
            <xsl:with-param name="sheets"><xsl:copy-of select="$sheetContents" /></xsl:with-param>
            <xsl:with-param name="vbas"><xsl:copy-of select="$vbas" /></xsl:with-param>
        </xsl:call-template>
    </xsl:template>
</xsl:stylesheet>