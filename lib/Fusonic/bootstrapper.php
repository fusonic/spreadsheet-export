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

if(!defined("FUSONIC_BOOTSTRAPPER_EXECUTED"))
{
    // Define base path
    define("FUSONIC_DIR", __DIR__);

    // Register autoloader
    spl_autoload_register(function($name) {

        if(strpos($name, "Fusonic\\") === 0)
        {
            $className = str_replace("\\", DIRECTORY_SEPARATOR, $name);
            $path = FUSONIC_DIR . substr($className, 7) . ".php";
            if(!include_once($path))
            {
                throw new Exception("Could not load " . $name);
            }
        }

    });

    // Prevent bootstrapper from being executed twice
    define("FUSONIC_BOOTSTRAPPER_EXECUTED", true);
}