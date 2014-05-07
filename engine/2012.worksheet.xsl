<?xml version='1.0'?>
<xsl:stylesheet version="2.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">

    <!--
         The excel 2012 workbook file

        @author Pavel Ptacek
        @copyright Pavel Ptacek (c) 2012-2013
    -->

    <xsl:template name="generate_sheets">
        <xsl:param name="sheetNames" />
        <xsl:param name="sheets" />
        <xsl:param name="vbas" />

        <!-- Put the sheetX contents into correct places -->
        <xsl:for-each select="$sheets/*">
            <xsl:result-document href="xl/worksheets/sheet{position()}.xml">
                <xsl:copy-of select="." />
            </xsl:result-document>
        </xsl:for-each>

        <!-- Generate the workbook.xml -->
        <xsl:result-document href="xl/workbook.xml">
            <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">
              <fileVersion appName="xl" lastEdited="6" lowestEdited="4" rupBuild="14128"/>
              <workbookPr defaultThemeVersion="124226"/>
              <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
                <mc:Choice Requires="x15">
                  <x15ac:absPath url="C:\Users\Pavel Ptacek\Desktop\" xmlns:x15ac="http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac"/>
                </mc:Choice>
              </mc:AlternateContent>
              <bookViews>
                <workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="12435"/>
              </bookViews>
              <sheets>
                <xsl:for-each select="$sheetNames/name">
                    <sheet name="{.}" sheetId="{position()}" r:id="rId{position()}"/>
                </xsl:for-each>
              </sheets>
              <calcPr calcId="125725"/>
            </workbook>
        </xsl:result-document>

        <!-- Generate the vba project files -->
        <xsl:if test="$vbas">
            <xsl:for-each select="$vbas/vba">
                <xsl:result-document href="xl/{.}" media-type="text/plain" omit-xml-declaration="yes">
                    <xsl:fallback/>
                </xsl:result-document>
            </xsl:for-each>
        </xsl:if>

        <!-- Generate the empty sharedStrings file -->
        <xsl:result-document href="xl/sharedStrings.xml">
            <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0" />
        </xsl:result-document>

    </xsl:template>

</xsl:stylesheet>