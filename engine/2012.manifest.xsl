<?xml version="1.0"?>
<xsl:stylesheet version="2.0" xmlns="http://schemas.openxmlformats.org/package/2006/content-types" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">

    <!--
         The excel 2012 manifest file

         @author Pavel Ptacek
         @copyright Pavel Ptacek (c) 2012-2013
    -->

    <!-- Generator macro -->
    <xsl:template name="generate_manifest">
        <xsl:param name="sheets" />
        <xsl:param name="drawings" />
        <xsl:param name="media" />
        <xsl:param name="vbas" />

        <!-- Standard manifest -->
        <xsl:result-document href="{encode-for-uri('[Content_Types].xml')}">
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                <xsl:if test="$media">
                <xsl:for-each select="$media/image">
                    <Default>
                        <xsl:attribute name="Extension" select="./ext" />
                        <xsl:attribute name="ContentType" select="./mime" />
                    </Default>
                </xsl:for-each>
                </xsl:if>
                <xsl:call-template name="man_defaults" />
                <xsl:call-template name="man_workbook"><xsl:with-param name="vbas"><xsl:copy-of select="$vbas" /></xsl:with-param></xsl:call-template>
                <xsl:call-template name="man_sheet_generator"><xsl:with-param name="sheets" select="$sheets" /></xsl:call-template>
                <xsl:call-template name="man_theme" />
                <xsl:call-template name="man_styles" />
                <xsl:call-template name="man_sharedStrings" />
                <xsl:call-template name="man_drawing_generator"><xsl:with-param name="drawings" select="$drawings" /></xsl:call-template>
                <xsl:if test="$vbas">
                <xsl:for-each select="$vbas/vba">
                    <Override PartName="/xl/{current()}" ContentType="application/vnd.ms-office.vbaProject"/>
                </xsl:for-each>
                </xsl:if>
                <xsl:call-template name="man_endings" />
            </Types>
        </xsl:result-document>

        <!-- The main rels file -->
        <xsl:result-document href="./_rels/\.rels">
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
              <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
            </Relationships>
        </xsl:result-document>

        <!-- The workbook rels file -->
        <xsl:result-document href="xl/_rels/workbook.xml.rels">
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <xsl:call-template name="man_wkrels_generator">
                    <xsl:with-param name="sheets" select="$sheets" />
                </xsl:call-template>
                <Relationship Id="rId{$sheets + 3}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
                <Relationship Id="rId{$sheets + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
                <Relationship Id="rId{$sheets + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
                <xsl:if test="$vbas">
                <xsl:for-each select="$vbas/vba">
                    <Relationship Id="rId{$sheets + 4 + position()}" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="{.}"/>
                </xsl:for-each>
                </xsl:if>
            </Relationships>
        </xsl:result-document>
    </xsl:template>


    <!-- The core macro -->
    <xsl:template name="man_add">
        <xsl:param name="path" />
        <xsl:param name="type" />
        <Override PartName="{$path}" ContentType="{$type}" />
    </xsl:template>

    <!-- default sheets -->
    <xsl:template name="man_defaults">
        <Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"/>
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
    </xsl:template>

    <!-- Individual file parameters -->
    <xsl:template name="man_workbook">
        <xsl:param name="vbas" />

        <xsl:call-template name="man_add">
            <xsl:with-param name="path">/xl/workbook.xml</xsl:with-param>
            <xsl:with-param name="type">
                <xsl:choose>
                    <xsl:when test="$vbas/vba">application/vnd.ms-excel.sheet.macroEnabled.main+xml</xsl:when>
                    <xsl:otherwise>application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml</xsl:otherwise>
                </xsl:choose>
            </xsl:with-param>
        </xsl:call-template>
    </xsl:template>

    <!-- one sheet -->
    <xsl:template name="man_sheet">
        <xsl:param name="nr" />

        <xsl:variable name="path">
            <xsl:text>/xl/worksheets/sheet</xsl:text>
            <xsl:value-of select="$nr" />
            <xsl:text>.xml</xsl:text>
        </xsl:variable>

        <xsl:call-template name="man_add">
            <xsl:with-param name="path"><xsl:value-of select="$path" /></xsl:with-param>
            <xsl:with-param name="type">application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml</xsl:with-param>
        </xsl:call-template>
    </xsl:template>

    <!-- sheet generator helper -->
    <xsl:template name="man_sheet_generator">
        <xsl:param name="sheets" />

        <xsl:if test="number($sheets) > 0">
            <!-- at first, recurse down -->
            <xsl:call-template name="man_sheet_generator">
                <xsl:with-param name="sheets" select="$sheets - 1" />
            </xsl:call-template>

            <!-- Output the sheet -->
            <xsl:call-template name="man_sheet">
                <xsl:with-param name="nr" select="$sheets" />
            </xsl:call-template>
        </xsl:if>
    </xsl:template>

    <xsl:template name="man_theme">
        <xsl:call-template name="man_add">
            <xsl:with-param name="path">/xl/theme/theme1.xml</xsl:with-param>
            <xsl:with-param name="type">application/vnd.openxmlformats-officedocument.theme+xml</xsl:with-param>
        </xsl:call-template>
    </xsl:template>

    <xsl:template name="man_styles">
        <xsl:call-template name="man_add">
            <xsl:with-param name="path">/xl/styles.xml</xsl:with-param>
            <xsl:with-param name="type">application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml</xsl:with-param>
        </xsl:call-template>
    </xsl:template>

    <xsl:template name="man_sharedStrings">
        <xsl:call-template name="man_add">
            <xsl:with-param name="path">/xl/sharedStrings.xml</xsl:with-param>
            <xsl:with-param name="type">application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml</xsl:with-param>
        </xsl:call-template>
    </xsl:template>

    <xsl:template name="man_drawing">
        <xsl:param name="nr" />

        <xsl:variable name="path">
            <xsl:text>/xl/drawings/drawing</xsl:text>
            <xsl:value-of select="$nr" />
            <xsl:text>.xml</xsl:text>
        </xsl:variable>

        <xsl:call-template name="man_add">
            <xsl:with-param name="path"><xsl:value-of select="$path" /></xsl:with-param>
            <xsl:with-param name="type">application/vnd.openxmlformats-officedocument.drawing+xml</xsl:with-param>
        </xsl:call-template>
    </xsl:template>

    <!-- the drawing helper -->
    <xsl:template name="man_drawing_generator">
        <xsl:param name="drawings" />

        <xsl:if test="number($drawings) > 0">
            <!-- at first, recurse down -->
            <xsl:call-template name="man_drawing_generator">
                <xsl:with-param name="drawings" select="$drawings - 1" />
            </xsl:call-template>

            <!-- Output the sheet -->
            <xsl:call-template name="man_drawing">
                <xsl:with-param name="nr" select="$drawings" />
            </xsl:call-template>
        </xsl:if>
    </xsl:template>

    <xsl:template name="man_endings">
        <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
        <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
    </xsl:template>

    <!-- The rels generator -->
    <xsl:template name="man_wkrels_generator" xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <xsl:param name="sheets" />

        <xsl:if test="number($sheets) > 0">
            <!-- first, recurse down -->
            <xsl:call-template name="man_wkrels_generator">
                <xsl:with-param name="sheets" select="$sheets - 1" />
            </xsl:call-template>

            <!-- And output the current sheet line -->
            <xsl:variable name="rId">
                <xsl:text>rId</xsl:text>
                <xsl:value-of select="$sheets" />
            </xsl:variable>
            <xsl:variable name="sheetName">
                <xsl:text>worksheets/sheet</xsl:text>
                <xsl:value-of select="$sheets" />
                <xsl:text>.xml</xsl:text>
            </xsl:variable>

            <Relationship Id="{$rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="{$sheetName}"/>
        </xsl:if>
    </xsl:template>
</xsl:stylesheet>