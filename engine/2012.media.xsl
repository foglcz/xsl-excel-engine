<?xml version='1.0'?>
<xsl:stylesheet version="2.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">

    <!--
        Excel 2012 media file

        @author Pavel Ptacek
        @copyright Pavel Ptacek (c) 2012-2013
    -->

    <!-- Main generator -->
    <xsl:template name="generate_media">
        <xsl:param name="images" />
        <xsl:param name="printerSettings" />

        <!-- Generate the drawing for every sheet -->
        <xsl:if test="$images">
            <xsl:for-each-group select="$images/image" group-by="./child::*/sheet/@nr">
                <xsl:variable name="sheetNr"><xsl:value-of select="current-grouping-key()" /></xsl:variable>
                <xsl:call-template name="gen_media_drawings">
                    <xsl:with-param name="sheetNr" select="current-grouping-key()" />
                    <xsl:with-param name="images" select="current-group()" />
                    <xsl:with-param name="printerSettings" select="$printerSettings/setting[@for=$sheetNr]" />
                </xsl:call-template>
            </xsl:for-each-group>
        </xsl:if>
        <xsl:if test="not($images) and $printerSettings">
            <xsl:for-each select="$printerSettings/setting">
                <xsl:call-template name="gen_media_drawings">
                    <xsl:with-param name="printerSettings" select="." />
                    <xsl:with-param name="sheetNr" select="@for" />
                </xsl:call-template>
            </xsl:for-each>
        </xsl:if>

        <!-- Binaries - put the files into correct places -->
        <xsl:for-each-group select="$images/image" group-by="concat(path, '.', ext)">
            <xsl:result-document href="xl/media/{current-grouping-key()}" media-type="text/plain" omit-xml-declaration="yes">
                <xsl:fallback/>
            </xsl:result-document>
        </xsl:for-each-group>
        <xsl:if test="$printerSettings">
            <xsl:for-each select="$printerSettings/setting">
                <xsl:result-document href="xl/printerSettings/{.}" media-type="text/plain" omit-xml-declaration="yes">
                    <xsl:fallback />
                </xsl:result-document>
            </xsl:for-each>
        </xsl:if>
    </xsl:template>

    <!-- This template takes the images for the given sheet, generates the drawing.xml, drawing.xml.rels, sheetX.xml.rels -->
    <xsl:template name="gen_media_drawings">
        <xsl:param name="sheetNr" />
        <xsl:param name="images" />
        <xsl:param name="printerSettings" />

        <!-- Generate the sheetX.xml.rels file -->
        <xsl:result-document href="xl/worksheets/_rels/sheet{$sheetNr}.xml.rels">
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <xsl:if test="$images">
                    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing{$sheetNr}.xml"/>
                </xsl:if>
                <xsl:if test="$printerSettings">
                    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings" Target="../printerSettings/{$printerSettings}"/>
                </xsl:if>
            </Relationships>
        </xsl:result-document>

        <!-- Generate the drawing & rels -->
        <xsl:if test="$images">
            <!-- Generate the drawingX.xml file -->
            <xsl:result-document href="xl/drawings/drawing{$sheetNr}.xml">
                <xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                    <xsl:for-each select="$images">
                        <xsl:variable name="imgId" select="position()" />

                        <xsl:for-each select="./xdr:from">
                          <xdr:twoCellAnchor editAs="oneCell">
                            <xsl:copy-of select="../xdr:from[position()]" />
                            <xsl:copy-of select="../xdr:to[position()]" />
                            <xdr:pic>
                              <xdr:nvPicPr>
                                <xdr:cNvPr id="{$imgId}" name="Picture {$imgId}"/>
                                <xdr:cNvPicPr>
                                  <a:picLocks noChangeAspect="1"/>
                                </xdr:cNvPicPr>
                              </xdr:nvPicPr>
                              <xdr:blipFill>
                                <a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId{$imgId}" cstate="print">
                                  <a:extLst>
                                    <a:ext>
                                        <xsl:attribute name="uri">
                                            <xsl:text>{</xsl:text>
                                            <xsl:text>28A0092B-C50C-407E-A947-70E740481C1C</xsl:text>
                                            <xsl:text>}</xsl:text>
                                        </xsl:attribute>
                                        <a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/>
                                    </a:ext>
                                  </a:extLst>
                                </a:blip>
                                <a:stretch>
                                  <a:fillRect/>
                                </a:stretch>
                              </xdr:blipFill>
                              <xsl:copy-of select="../xdr:spPr[position()]" />
                            </xdr:pic>
                            <xdr:clientData/>
                          </xdr:twoCellAnchor>
                        </xsl:for-each>
                    </xsl:for-each>
                </xdr:wsDr>
            </xsl:result-document>

            <!-- Generate the drawingX.xml.rels file -->
            <xsl:result-document href="xl/drawings/_rels/drawing{$sheetNr}.xml.rels">
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                    <xsl:for-each select="$images">
                      <Relationship Id="rId{position()}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/{./path}.{./ext}"/>
                    </xsl:for-each>
                </Relationships>
            </xsl:result-document>
        </xsl:if>
    </xsl:template>


</xsl:stylesheet>