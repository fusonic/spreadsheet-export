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

class CsvWriter extends Writer
{
    const READ_CHUNK_SIZE = 1024;
    const CHARSET_UTF8 = 1;
    const CHARSET_ISO = 2;

    public $delimiter = ",";
    public $enclosure = "\"";
    public $charset = self::CHARSET_UTF8;

    public function GetContentType()
    {
        return "text/csv; charset=" . ($this->charset == self::CHARSET_UTF8 ? "UTF-8" : "ISO-8859-1");
    }

    public function GetDefaultExtension()
    {
        return "csv";
    }

    public function GetContent(array $columns, array $data)
    {
        // Create a temporary filestream to use PHP CSV methods
        $fd = fopen("php://temp", "r+");

        // Write content
        foreach($data as $row)
        {
            if(!is_array($row))
            {
                throw new Exception("Row is not an array.");
            }

            fputcsv($fd, $row, $this->delimiter, $this->enclosure);
        }

        // Read content
        rewind($fd);
        $content = "";
        while($chunk = fread($fd, self::READ_CHUNK_SIZE))
        {
            $content .= $chunk;
        }

        // Clean up
        fclose($fd);

        // Return correctly encoded content
        switch($this->charset)
        {
            case self::CHARSET_ISO:
                return utf8_decode($content);

            default:
                return $content;
        }
    }
}
