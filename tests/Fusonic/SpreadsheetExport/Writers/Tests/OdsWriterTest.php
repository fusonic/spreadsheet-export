<?php

use Fusonic\SpreadsheetExport\ColumnTypes\TextColumn;
use Fusonic\SpreadsheetExport\Spreadsheet;
use Fusonic\SpreadsheetExport\Writers\OdsWriter;

class OdsWriterTest extends PHPUnit_Framework_TestCase
{
    private $writer;

    public function setUp()
    {
        parent::setUp();

        $this->writer = new OdsWriter();
    }

    public function testIssue4_usingAmpersandInTextCrashes()
    {
        $sheet = new Spreadsheet();
        $sheet->addColumn(new TextColumn("Title"));
        $sheet->addRow(array("The good, the bad & the ugly"));
        $sheet->get($this->writer);
    }
}
