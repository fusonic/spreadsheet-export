<?php

/*
 * Copyright (c) 2012 Fusonic GmbH
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 */

namespace Fusonic\SpreadsheetExport\Writers;
use Fusonic\SpreadsheetExport\Writer;
use Fusonic\SpreadsheetExport\ColumnTypes\CurrencyColumn;
use Fusonic\SpreadsheetExport\ColumnTypes\DateColumn;
use Fusonic\SpreadsheetExport\ColumnTypes\NumericColumn;
use Fusonic\SpreadsheetExport\ColumnTypes\TextColumn;

class OdsWriter extends Writer
{
    const ODF_NAMESPACE_MANIFEST = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";
    const ODF_NAMESPACE_OFFICE = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
    const ODF_NAMESPACE_STYLE = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
    const ODF_NAMESPACE_TEXT = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
    const ODF_NAMESPACE_TABLE = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
    const ODF_VERSION = "1.1";

    public function __construct()
    {
        if(!class_exists("ZipArchive", false))
        {
            throw new \RuntimeException("Ods writer requires zip extension to be installed.");
        }
    }

    public function GetContentType()
    {
        return "application/vnd.oasis.opendocument.spreadsheet";
    }

    public function GetDefaultExtension()
    {
        return "ods";
    }

    public function GetContent(array $columns, array $data)
    {
        $tmpName = tempnam(sys_get_temp_dir(), "SpreadsheetWriterOds");

        // Open zip file
        $zip = new \ZipArchive();
        $zip->open($tmpName, \ZipArchive::CREATE | \ZipArchive::OVERWRITE);

        // Set mime type
        $zip->addFromString("mimetype", "application/vnd.oasis.opendocument.spreadsheet");

        // Write meta information (manifest)
        $xml = new \SimpleXMLElement('<manifest:manifest '
            . 'xmlns:manifest="' . self::ODF_NAMESPACE_MANIFEST . '" />');

        $xmlFileEntry = $xml->addChild("file-entry", null, self::ODF_NAMESPACE_MANIFEST);
        $xmlFileEntry->addAttribute("manifest:media-type", "application/vnd.oasis.opendocument.spreadsheet", self::ODF_NAMESPACE_MANIFEST);
        $xmlFileEntry->addAttribute("manifest:full-path", "/", self::ODF_NAMESPACE_MANIFEST);

        $xmlFileEntry = $xml->addChild("file-entry", null, self::ODF_NAMESPACE_MANIFEST);
        $xmlFileEntry->addAttribute("manifest:media-type", "text/xml", self::ODF_NAMESPACE_MANIFEST);
        $xmlFileEntry->addAttribute("manifest:full-path", "content.xml", self::ODF_NAMESPACE_MANIFEST);

        $zip->addEmptyDir("META-INF");
        $zip->addFromString("META-INF/manifest.xml", $xml->asXML());

        // Write content (content.xml)
        $xml = new \SimpleXMLElement('<office:document-content '
            . 'xmlns:office="' . self::ODF_NAMESPACE_OFFICE . '" '
            . 'xmlns:style="' . self::ODF_NAMESPACE_STYLE . '" '
            . 'xmlns:text="' . self::ODF_NAMESPACE_TEXT . '" '
            . 'xmlns:table="' . self::ODF_NAMESPACE_TABLE . '" '
            . 'office:version="1.1" />');


        // Set styles
        $xmlAutomaticStyles = $xml->addChild("automatic-styles", null, self::ODF_NAMESPACE_OFFICE);
        foreach($columns AS $columnIndex => $column)
        {
            // <style>
            $xmlStyle = $xmlAutomaticStyles->addChild("style", null, self::ODF_NAMESPACE_STYLE);
            $xmlStyle->addAttribute("style:name", "col" . $columnIndex, self::ODF_NAMESPACE_STYLE);
            $xmlStyle->addAttribute("style:family", "table-column", self::ODF_NAMESPACE_STYLE);

            // <table-coluömn-properties>
            $xmlTableColumnProperties = $xmlStyle->addChild("table-column-properties", null, self::ODF_NAMESPACE_STYLE);
            $xmlTableColumnProperties->addAttribute("style:column-width", $column->width . "cm", self::ODF_NAMESPACE_STYLE);
        }

        // Write table
        $xmlBody = $xml->addChild("body", null, self::ODF_NAMESPACE_OFFICE);
        $xmlSpreadsheet = $xmlBody->addChild("spreadsheet", null, self::ODF_NAMESPACE_OFFICE);
        $xmlTable = $xmlSpreadsheet->addChild("table", null, self::ODF_NAMESPACE_TABLE);

        // Columns
        foreach($columns AS $columnIndex => $column)
        {
            // <table-column>
            $xmlTableColumn = $xmlTable->addChild("table-column", null, self::ODF_NAMESPACE_TABLE);
            $xmlTableColumn->addAttribute("table:style-name", "col" . $columnIndex, self::ODF_NAMESPACE_TABLE);
        }

        // Rows
        foreach($data AS $rowIndex => $row)
        {
            $xmlRow = $xmlTable->addChild("table-row", null, self::ODF_NAMESPACE_TABLE);

            // Cells
            foreach($columns AS $columnIndex => $column) {

                // <table-cell>
                $xmlCell = $xmlRow->addChild("table-cell", null, self::ODF_NAMESPACE_TABLE);

                if($column instanceof TextColumn)
                {
                    $xmlCell->addAttribute("office:value-type", "string", self::ODF_NAMESPACE_OFFICE);
                    $xmlCell->addChild("p", (string)$row[$columnIndex], self::ODF_NAMESPACE_TEXT);
                }
                elseif($column instanceof CurrencyColumn)
                {
                    $xmlCell->addAttribute("office:value-type", "currency", self::ODF_NAMESPACE_OFFICE);
                    $xmlCell->addAttribute("office:currency", strtoupper($column->currency), self::ODF_NAMESPACE_OFFICE);
                    $xmlCell->addAttribute("office:value", (float)$row[$columnIndex], self::ODF_NAMESPACE_OFFICE);
                    $xmlCell->addChild("p", (float)$row[$columnIndex], self::ODF_NAMESPACE_TEXT);
                }
                elseif($column instanceof NumericColumn)
                {
                    $xmlCell->addAttribute("office:value-type", "float", self::ODF_NAMESPACE_OFFICE);
                    $xmlCell->addAttribute("office:value", (float)$row[$columnIndex], self::ODF_NAMESPACE_OFFICE);
                    $xmlCell->addChild("p", (float)$row[$columnIndex], self::ODF_NAMESPACE_TEXT);
                }
                elseif($column instanceof DateColumn)
                {
                    $xmlCell->addAttribute("office:value-type", "date", self::ODF_NAMESPACE_OFFICE);

                    $value = $row[$columnIndex];
                    if(!$value instanceof DateTime)
                    {
                        $value = new \DateTime($value);
                    }

                    $xmlCell->addAttribute("office:date-value", substr($value->format("c"), 0, 19), self::ODF_NAMESPACE_OFFICE);
                    $xmlCell->addChild("p", $row[$columnIndex], self::ODF_NAMESPACE_TEXT);
                }
            }
        }

        $zip->addFromString("content.xml", $xml->asXML());

        // Close zip
        $zip->close();

        // Read content
        $content = file_get_contents($tmpName);

        // Cleanup
        @unlink($tmpName);

        return $content;
    }
}
