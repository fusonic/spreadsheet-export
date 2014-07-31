---
layout: page
title: Examples
permalink: /examples/
---

## Exporting a spreadsheet

{% highlight php startinline %}
use Fusonic\SpreadsheetExport\Spreadsheet;
use Fusonic\SpreadsheetExport\ColumnTypes\CurrencyColumn;
use Fusonic\SpreadsheetExport\ColumnTypes\DateColumn;
use Fusonic\SpreadsheetExport\ColumnTypes\NumericColumn;
use Fusonic\SpreadsheetExport\ColumnTypes\TextColumn;
use Fusonic\SpreadsheetExport\Writers\CsvWriter;
use Fusonic\SpreadsheetExport\Writers\TsvWriter;
use Fusonic\SpreadsheetExport\Writers\OdsWriter;

// Instantiate new spreadsheet
$export = new Spreadsheet();

// Add columns
$export->addColumn(new DateColumn("Date"));
$export->addColumn(new TextColumn("Comment"));
$export->addColumn(new NumericColumn("Population"));
$bipCol = new CurrencyColumn("GWP");
$bipCol->currency = CurrencyColumn::CURRENCY_USD;
$export->addColumn($bipCol);

// Add data rows
$export->addRow(array("1987-01-01", "world population reached 5 billion", 5, 24000000000000));
$export->addRow(array("1999-01-01", "world population reached 6 billion", 6, 41000000000000));
$export->addRow(array("2012-01-01", "world population reaches 7 billion", 7, null));

// Instantiate writer (CSV)
// $writer = new CsvWriter();
// $writer->includeColumnHeaders = true;
// $writer->charset = CsvWriter::CHARSET_ISO;

// Instantiate writer (TSV)
// $writer = new TsvWriter();
// $writer->includeColumnHeaders = true;
// $writer->charset = TsvWriter::CHARSET_ISO;

// Instantiate writer (ODS)
$writer = new OdsWriter();
$writer->includeColumnHeaders = true;

// Save
// $export->save($writer, "/tmp/Sample");

// Download
$export->download($writer, "Sample");
{% endhighlight %}