<xsl:stylesheet version="2.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
    <xsl:include href="../../engine/2012.engine.xsl" />
    <xsl:include href="static.xslt" />

    <!-- put your data-preprocessing in here -->

    <xsl:variable name="images">
        <image>
            <path>foglcz</path>
            <ext>png</ext>
            <mime>image/png</mime>
            <usedIn>
                <sheet nr="1" />
            </usedIn>

            <!-- following variables has been taken from drawing1.xml -->
            <xdr:from>
                <xdr:col>0</xdr:col>
                <xdr:colOff>85725</xdr:colOff>
                <xdr:row>0</xdr:row>
                <xdr:rowOff>104774</xdr:rowOff>
            </xdr:from>
            <xdr:to>
                <xdr:col>2</xdr:col>
                <xdr:colOff>895350</xdr:colOff>
                <xdr:row>6</xdr:row>
                <xdr:rowOff>81163</xdr:rowOff>
            </xdr:to>
            <xdr:spPr>
                <a:xfrm>
                    <a:off x="85725" y="104774"/>
                    <a:ext cx="1695450" cy="1119389"/>
                </a:xfrm>
                <a:prstGeom prst="rect">
                    <a:avLst/>
                </a:prstGeom>
            </xdr:spPr>
        </image>
    </xsl:variable>

    <xsl:variable name="contents">
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
            <sheetPr codeName="Sheet1"/>
            <dimension ref="">
                <xsl:attribute name="ref">
                    <xsl:text>A1:F</xsl:text>
                    <xsl:value-of select="count(data/data) + 7" />
                </xsl:attribute>
            </dimension>
            <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0">
                    <selection activeCell="K10" sqref="K10"/>
                </sheetView>
            </sheetViews>
            <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
            <cols>
                <col min="1" max="1" width="3" bestFit="1" customWidth="1"/>
                <col min="2" max="2" width="10.28515625" customWidth="1"/>
                <col min="3" max="3" width="16.42578125" customWidth="1"/>
                <col min="4" max="4" width="52.28515625" customWidth="1"/>
                <col min="5" max="5" width="18.5703125" customWidth="1"/>
                <col min="6" max="6" width="29.7109375" customWidth="1"/>
            </cols>
            <sheetData>
                <!-- header -->
                <row r="1" spans="1:6" x14ac:dyDescent="0.25">
                    <c r="D1" t="inlineStr">
                        <is><t>The FULL example of the generator</t></is>
                    </c>
                </row>
                <row r="3" spans="1:6" x14ac:dyDescent="0.25">
                    <c r="D3" t="inlineStr">
                        <is><t>Generates file with image, vbaMacro &amp; barcodes.</t></is>
                    </c>
                </row>
                <row r="4" spans="1:6" x14ac:dyDescent="0.25">
                    <c r="D4" t="inlineStr">
                        <is><t>In order for the barcodes to work, you need to have code128.ttf installed.</t></is>
                    </c>
                </row>
                <row r="6" spans="1:6" x14ac:dyDescent="0.25">
                    <c r="D6" t="inlineStr">
                        <is><t>Author is Pavel Ptacek</t></is>
                    </c>
                </row>
                <row r="7" spans="1:6" ht="15.75" thickBot="1" x14ac:dyDescent="0.3"/>

                <!-- The data table: first row (header) -->
                <row r="8" spans="1:6" ht="15.75" thickBot="1" x14ac:dyDescent="0.3">
                    <c r="A8" s="1" t="inlineStr">
                        <is><t>#</t></is>
                    </c>
                    <c r="B8" s="2" t="inlineStr">
                        <is><t>Date</t></is>
                    </c>
                    <c r="C8" s="18" t="inlineStr">
                        <is><t>Name</t></is>
                    </c>
                    <c r="D8" s="20"/>
                    <c r="E8" s="18" t="inlineStr">
                        <is><t>Barcode</t></is>
                    </c>
                    <c r="F8" s="19"/>
                </row>

                <xsl:for-each select="data/data">
                    <xsl:variable name="rowNr" select="position() + 8" />
                    <xsl:choose>
                        <!-- The data table: loop actual data -->
                        <xsl:when test="position() lt count(/data/data)">
                            <row r="{$rowNr}" spans="1:6" ht="16.5" x14ac:dyDescent="0.25">
                                <c r="A{$rowNr}" s="5">
                                    <v><xsl:value-of select="position()" /></v>
                                </c>
                                <c r="B{$rowNr}" s="9" t="e">
                                    <f>
                                        <xsl:text>DATEVALUE("</xsl:text>
                                        <xsl:value-of select="date" />
                                        <xsl:text>")</xsl:text>
                                    </f>
                                </c>
                                <c r="C{$rowNr}" s="24" t="inlineStr">
                                    <is><t><xsl:value-of select="name" /></t></is>
                                </c>
                                <c r="D{$rowNr}" s="24"/>
                                <c r="E{$rowNr}" s="3" t="inlineStr">
                                    <is><t><xsl:value-of select="barcode" /></t></is>
                                </c>
                                <c r="F{$rowNr}" s="21" t="e">
                                    <f>
                                        <xsl:text>Code128(E</xsl:text>
                                        <xsl:value-of select="$rowNr" />
                                        <xsl:text>)</xsl:text>
                                    </f>
                                </c>
                            </row>
                        </xsl:when>

                        <!-- Last row contains different style numbers due to borders -->
                        <xsl:otherwise>
                            <row r="{$rowNr}" spans="1:6" ht="17.25" thickBot="1" x14ac:dyDescent="0.3">
                                <c r="A{$rowNr}" s="7">
                                    <v><xsl:value-of select="position()" /></v>
                                </c>
                                <c r="B{$rowNr}" s="11" t="e">
                                    <f>
                                        <xsl:text>DATEVALUE("</xsl:text>
                                        <xsl:value-of select="date" />
                                        <xsl:text>")</xsl:text>
                                    </f>
                                </c>
                                <c r="C{$rowNr}" s="16" t="inlineStr">
                                    <is><t><xsl:value-of select="name" /></t></is>
                                </c>
                                <c r="D{$rowNr}" s="17"/>
                                <c r="E{$rowNr}" s="8" t="inlineStr">
                                    <is><t><xsl:value-of select="barcode" /></t></is>
                                </c>
                                <c r="F{$rowNr}" s="23" t="e">
                                    <f>
                                        <xsl:text>Code128(E</xsl:text>
                                        <xsl:value-of select="$rowNr" />
                                        <xsl:text>)</xsl:text>
                                    </f>
                                </c>
                            </row>
                        </xsl:otherwise>
                    </xsl:choose>
                </xsl:for-each>
            </sheetData>
            <mergeCells count="{count(data/data)}">
                <xsl:for-each select="data/data">
                    <mergeCell ref="C{position() + 8}:D{position() + 8}"/>
                </xsl:for-each>
            </mergeCells>
            <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
            <pageSetup paperSize="9" orientation="landscape" r:id="rId2"/>
            <drawing r:id="rId1"/>
        </worksheet>
    </xsl:variable>

    <xsl:template match="/">
        <xsl:call-template name="generate_excel">
            <xsl:with-param name="author">ACME Corp.</xsl:with-param>
            <xsl:with-param name="themes"><xsl:copy-of select="$themes" /></xsl:with-param>
            <xsl:with-param name="styles"><xsl:copy-of select="$styles" /></xsl:with-param>
            <xsl:with-param name="images"><xsl:copy-of select="$images" /></xsl:with-param>
            <xsl:with-param name="sheetContents"><xsl:copy-of select="$contents" /></xsl:with-param>
            <xsl:with-param name="printerSettings"><setting for="1">Example_printer.bin</setting></xsl:with-param>
            <xsl:with-param name="vbas"><vba>Example_Vba.bin</vba></xsl:with-param>
        </xsl:call-template>
    </xsl:template>
</xsl:stylesheet>