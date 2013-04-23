<?php
/**
 * This file is part of the foglcz/xsl-excel-engine github project.
 *
 * How to use:
 * 1) extract your .xslx file as zip into folder
 * 2) pick your sheetX.xml file & sharedStrings.xml file -> and put them into the directory as this script.
 * 3) rename sheetX.xml into sheet.xml
 * 3) $ php sharedStringsConvertor.php
 * 4) open "parsedSheet.xml" file in your favourite XML editor.
 *
 * @author Pavel Ptacek
 * @copyright Pavel Ptacek (c) 2012-2013
 */
if(!file_exists('sheet.xml')) {
    echo 'Please make sure that sheet.xml file exists.';
    exit;
}
if(!file_exists('sharedStrings.xml')) {
    echo 'Please make sure that sharedStrings.xml file exists';
    exit;
}
if(!function_exists('simplexml_load_file')) {
    echo 'Please make sure you have SimpleXML php extension enabled.';
    exit;
}

/**
 * The magic happens below.
 */
// load both
$sheet = simplexml_load_file('sheet.xml');
$strings = simplexml_load_file('sharedStrings.xml');

// Loop sheet & update
foreach($sheet->sheetData->row as $row) {
    foreach($row->c as $cell) {
        $attrs = $cell->attributes();

        // Check for string presence & get the string
        if(isset($attrs['t']) && $attrs['t'] == 's') {
            $string = (string)$strings->si[intval((string)$cell->v[0])]->t;

            // revamp from reference to sharedStrings.xml into inline string support.
            unset($cell->v);
            $attrs['t'] = 'inlineStr';
            $is = $cell->addChild('is');
            $is->addChild('t');
            $is->t = $string;
        }
    }
}

// done & done
$sheet->saveXML('parsedSheet.xml');
echo 'Your parsedSheet.xml is available for use.';
exit;
