<?php

/*
 * Copyright (c) 2012-2013 Fusonic GmbH
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

namespace Fusonic\SpreadsheetExport;

class Spreadsheet
{

	private $columns = array();
	private $data = array();

    public $appendDefaultExtension = true;

	public function addColumn(Column $column)
    {
        $this->columns[] = $column;
	}

	public function addColumns(Column $column, $amount)
    {
        for($i = 0; $i < $amount; $i++)
        {
            $this->addColumn($column);
        }
	}

	public function addRow(array $data)
    {
		$this->data[] = $data;
	}

	private function sendGeneralHeaders()
    {
		header("Cache-Control: must-revalidate, post-check=0, pre-check=0");
		header("Pragma: public"); // For Internet Explorer
	}

    public function get(Writer $writer)
    {
        return $writer->getContent($this->columns, $this->data);
    }

    public function download(Writer $writer, $filename = null)
    {
        $content = $this->get($writer);

        // Send headers
        $this->sendGeneralHeaders();
        header("Content-Type: " . $writer->getContentType());
        header("Content-Length: " . strlen($content));
        if($filename !== null)
        {
            $extension = pathinfo($filename, PATHINFO_EXTENSION);
            if(!$extension && $this->appendDefaultExtension)
            {
                $filename .= "." . $writer->getDefaultExtension();
            }

            header("Content-Disposition: attachment; filename=\"" . $filename . "\"");
        }

        echo $content;
    }

    public function save(Writer $writer, $path)
    {
        $extension = pathinfo($path, PATHINFO_EXTENSION);
        if(!$extension && $this->appendDefaultExtension)
        {
            $path .= "." . $writer->getDefaultExtension();
        }

        file_put_contents($path, $this->get($writer));
    }

}
