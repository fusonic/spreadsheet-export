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

namespace Fusonic\SpreadsheetExport;

class Spreadsheet
{

	private $columns = array();
	private $data = array();

    public $appendDefaultExtension = true;

	public function AddColumn(Column $column)
    {
        $this->columns[] = $column;
	}

	public function AddColumns(Column $column, $amount)
    {
        for($i = 0; $i < $amount; $i++)
        {
            $this->AddColumn($column);
        }
	}

	public function AddRow(array $data)
    {
		$this->data[] = $data;
	}

	private function SendGeneralHeaders()
    {
		header("Cache-Control: must-revalidate, post-check=0, pre-check=0");
		header("Pragma: public"); // For Internet Explorer
	}

    public function Get(Writer $writer)
    {
        return $writer->GetContent($this->columns, $this->data);
    }

    public function Download(Writer $writer, $filename = null)
    {
        $content = $this->Get($writer);

        // Send headers
        $this->SendGeneralHeaders();
        header("Content-Type: " . $writer->GetContentType());
        header("Content-Length: " . strlen($content));
        if($filename !== null)
        {
            $extension = pathinfo($filename, PATHINFO_EXTENSION);
            if(!$extension && $this->appendDefaultExtension)
            {
                $filename .= "." . $writer->GetDefaultExtension();
            }

            header("Content-Disposition: attachment; filename=\"" . $filename . "\"");
        }

        echo $content;
    }

    public function Save(Writer $writer, $path)
    {
        $extension = pathinfo($path, PATHINFO_EXTENSION);
        if(!$extension && $this->appendDefaultExtension)
        {
            $path .= "." . $writer->GetDefaultExtension();
        }

        file_put_contents($path, $this->Get($writer));
    }

}