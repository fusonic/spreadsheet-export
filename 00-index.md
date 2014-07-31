---
layout: index
title: Home
permalink: /
---

> SpreadsheetExport is a PHP library which allows you to export spreadsheet data in various formats while only writing code once. Currently supported are
>
> * OpenDocument Spreadsheet (.ods)
> * Comma Separated Values (.csv)
> * Tab Separated Values (.tsv)

## Requirements

* PHP 5.3 and up
* Zip-Extension to use ODS Export

## Installation

The most flexible installation method is using Composer: Simply create a composer.json file in the root of your project:

{% highlight json %}
{
    "require": {
        "fusonic/spreadsheetexport": "@dev"
    }
}
{% endhighlight %}

Install composer and run install command:

{% highlight bash %}
curl -s http://getcomposer.org/installer | php
php composer.phar install
{% endhighlight %}

Once installed, include vendor/autoload.php in your script.

{% highlight php startinline %}
require "vendor/autoload.php";
{% endhighlight %}

## License

This library is licensed under the MIT license.
