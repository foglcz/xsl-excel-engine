<?xml version='1.0'?>
<xsl:stylesheet version="2.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
    <xsl:include href="../../engine/2012.engine.xsl" />
    <xsl:include href="static.xsl" />

    <xsl:template match="/">
        <xsl:call-template name="generate_excel">
            <xsl:with-param name="sheetContents">
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">
                    <dimension ref="B1:L11"/>
                    <sheetViews>
                        <sheetView tabSelected="1" zoomScaleNormal="100" workbookViewId="0">
                            <selection activeCell="G3" sqref="G3"/>
                        </sheetView>
                    </sheetViews>
                    <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
                    <cols>
                        <col min="6" max="6" width="4.5703125" customWidth="1"/>
                        <col min="7" max="7" width="17.42578125" bestFit="1" customWidth="1"/>
                    </cols>
                    <sheetData>
                        <row r="1" spans="2:12" ht="36.75" thickBot="1" x14ac:dyDescent="0.6">
                            <c r="B1" s="5" t="inlineStr">

                                <is><t>foglcz/xsl-excel-engine</t></is></c>
                            <c r="C1" s="6"/>
                            <c r="D1" s="6"/>
                            <c r="E1" s="6"/>
                            <c r="F1" s="6"/>
                            <c r="G1" s="6"/>
                            <c r="H1" s="6"/>
                            <c r="I1" s="6"/>
                            <c r="J1" s="6"/>
                            <c r="K1" s="6"/>
                            <c r="L1" s="7"/>
                        </row>
                        <row r="2" spans="2:12" ht="15.75" thickBot="1" x14ac:dyDescent="0.3">
                            <c r="B2" s="1" t="inlineStr">

                                <is><t>image courtesy of the internetz</t></is></c>
                            <c r="C2" s="1"/>
                            <c r="D2" s="1"/>
                            <c r="E2" s="1"/>
                            <c r="G2" s="31" t="inlineStr">

                                <is><t>oh, by the way, try to print this sheet.</t></is></c>
                            <c r="H2" s="31"/>
                            <c r="I2" s="31"/>
                            <c r="J2" s="31"/>
                            <c r="K2" s="31"/>
                            <c r="L2" s="31"/>
                        </row>
                        <row r="3" spans="2:12" x14ac:dyDescent="0.25">
                            <c r="B3" s="12"/>
                            <c r="C3" s="13"/>
                            <c r="D3" s="13"/>
                            <c r="E3" s="14"/>
                            <c r="G3" s="8"/>
                            <c r="H3" s="19"/>
                            <c r="I3" s="19"/>
                            <c r="J3" s="19"/>
                            <c r="K3" s="19"/>
                            <c r="L3" s="20"/>
                        </row>
                        <row r="4" spans="2:12" x14ac:dyDescent="0.25">
                            <c r="B4" s="15"/>
                            <c r="C4" s="2"/>
                            <c r="D4" s="2"/>
                            <c r="E4" s="3"/>
                            <c r="G4" s="9" t="inlineStr">

                                <is><t>Brought to you by:</t></is></c>
                            <c r="H4" s="21" t="inlineStr">

                                <is><t>Pavel Ptacek</t></is></c>
                            <c r="I4" s="21"/>
                            <c r="J4" s="21"/>
                            <c r="K4" s="21"/>
                            <c r="L4" s="22"/>
                        </row>
                        <row r="5" spans="2:12" x14ac:dyDescent="0.25">
                            <c r="B5" s="15"/>
                            <c r="C5" s="2"/>
                            <c r="D5" s="2"/>
                            <c r="E5" s="3"/>
                            <c r="G5" s="9" t="inlineStr">

                                <is><t>Twitter handle:</t></is></c>
                            <c r="H5" s="23" t="inlineStr">

                                <is><t>@foglcz</t></is></c>
                            <c r="I5" s="23"/>
                            <c r="J5" s="23"/>
                            <c r="K5" s="23"/>
                            <c r="L5" s="24"/>
                        </row>
                        <row r="6" spans="2:12" x14ac:dyDescent="0.25">
                            <c r="B6" s="15"/>
                            <c r="C6" s="2"/>
                            <c r="D6" s="2"/>
                            <c r="E6" s="3"/>
                            <c r="G6" s="9" t="inlineStr">

                                <is><t>Linkedin:</t></is></c>
                            <c r="H6" s="25" t="inlineStr">

                                <is><t>/in/foglcz</t></is></c>
                            <c r="I6" s="25"/>
                            <c r="J6" s="25"/>
                            <c r="K6" s="25"/>
                            <c r="L6" s="26"/>
                        </row>
                        <row r="7" spans="2:12" x14ac:dyDescent="0.25">
                            <c r="B7" s="15"/>
                            <c r="C7" s="2"/>
                            <c r="D7" s="2"/>
                            <c r="E7" s="3"/>
                            <c r="G7" s="9"/>
                            <c r="H7" s="21"/>
                            <c r="I7" s="21"/>
                            <c r="J7" s="21"/>
                            <c r="K7" s="21"/>
                            <c r="L7" s="22"/>
                        </row>
                        <row r="8" spans="2:12" x14ac:dyDescent="0.25">
                            <c r="B8" s="15"/>
                            <c r="C8" s="2"/>
                            <c r="D8" s="2"/>
                            <c r="E8" s="3"/>
                            <c r="G8" s="9" t="inlineStr">

                                <is><t>Direct mail:</t></is></c>
                            <c r="H8" s="25" t="inlineStr">

                                <is><t>birdie@animalgroup.cz</t></is></c>
                            <c r="I8" s="25"/>
                            <c r="J8" s="25"/>
                            <c r="K8" s="25"/>
                            <c r="L8" s="26"/>
                        </row>
                        <row r="9" spans="2:12" x14ac:dyDescent="0.25">
                            <c r="B9" s="15"/>
                            <c r="C9" s="2"/>
                            <c r="D9" s="2"/>
                            <c r="E9" s="3"/>
                            <c r="G9" s="4"/>
                            <c r="H9" s="21"/>
                            <c r="I9" s="21"/>
                            <c r="J9" s="21"/>
                            <c r="K9" s="21"/>
                            <c r="L9" s="22"/>
                        </row>
                        <row r="10" spans="2:12" x14ac:dyDescent="0.25">
                            <c r="B10" s="15"/>
                            <c r="C10" s="2"/>
                            <c r="D10" s="2"/>
                            <c r="E10" s="3"/>
                            <c r="G10" s="10"/>
                            <c r="H10" s="27" t="inlineStr">

                                <is><t>Verba movent, exempla trahunt.</t></is></c>
                            <c r="I10" s="27"/>
                            <c r="J10" s="27"/>
                            <c r="K10" s="27"/>
                            <c r="L10" s="28"/>
                        </row>
                        <row r="11" spans="2:12" ht="15.75" thickBot="1" x14ac:dyDescent="0.3">
                            <c r="B11" s="16"/>
                            <c r="C11" s="17"/>
                            <c r="D11" s="17"/>
                            <c r="E11" s="18"/>
                            <c r="G11" s="11"/>
                            <c r="H11" s="29"/>
                            <c r="I11" s="29"/>
                            <c r="J11" s="29"/>
                            <c r="K11" s="29"/>
                            <c r="L11" s="30"/>
                        </row>
                    </sheetData>
                    <mergeCells count="13">
                        <mergeCell ref="H4:L4"/>
                        <mergeCell ref="H7:L7"/>
                        <mergeCell ref="H9:L9"/>
                        <mergeCell ref="H10:L10"/>
                        <mergeCell ref="H11:L11"/>
                        <mergeCell ref="B1:L1"/>
                        <mergeCell ref="H3:L3"/>
                        <mergeCell ref="H5:L5"/>
                        <mergeCell ref="H6:L6"/>
                        <mergeCell ref="H8:L8"/>
                        <mergeCell ref="B2:E2"/>
                        <mergeCell ref="B3:E11"/>
                        <mergeCell ref="G2:L2"/>
                    </mergeCells>
                    <printOptions horizontalCentered="1" verticalCentered="1"/>
                    <pageMargins left="0.70866141732283472" right="0.70866141732283472" top="0.74803149606299213" bottom="0.74803149606299213" header="0.31496062992125984" footer="0.31496062992125984"/>
                    <pageSetup paperSize="9" orientation="landscape" r:id="rId2"/>
                    <drawing r:id="rId1"/>
                </worksheet>
            </xsl:with-param>
        </xsl:call-template>
    </xsl:template>
</xsl:stylesheet>