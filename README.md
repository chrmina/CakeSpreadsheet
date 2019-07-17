[![Build Status](https://img.shields.io/travis/chrmina/CakeSpreadsheet/master.svg?style=flat-square)](https://travis-ci.org/chrmina/CakeSpreadsheet)
[![Coverage Status](https://img.shields.io/coveralls/chrmina/CakeSpreadsheet.svg?style=flat-square)](https://coveralls.io/r/chrmina/CakeSpreadsheet?branch=master)
[![Total Downloads](https://img.shields.io/packagist/dt/chrmina/CakeSpreadsheet.svg?style=flat-square)](https://packagist.org/packages/chrmina/CakeSpreadsheet)
[![Latest Stable Version](https://img.shields.io/packagist/v/chrmina/CakeSpreadsheet.svg?style=flat-square)](https://packagist.org/packages/chrmina/CakeSpreadsheet)

# CakeSpreadsheet

A plugin to generate Excel spreadsheet files with CakePHP. Modified from [CakeExcel](https://github.com/dakota/CakeExcel) to use PHPSpreadsheet instead of obsolete PHPExcel.

## Requirements

* CakePHP 3.x
* PHP 5.4.16 or greater
* Patience

## Installation

_[Using [Composer](http://getcomposer.org/)]_

```
composer require chrmina/CakeSpreadsheet
```

### Enable plugin

Load the plugin in your app's `config/bootstrap.php` file:

    Plugin::load('CakeSpreadsheet', ['bootstrap' => true, 'routes' => true]);

## Usage

First, you'll want to setup extension parsing for the `xlsx` extension. To do so, you will need to add the following to your `config/routes.php` file:

```php
# Set this before you specify any routes
Router::extensions('xlsx');
```

Next, we'll need to add a viewClassMap entry to your Controller. You can place the following in your AppController:

```php
$this->loadComponent('RequestHandler', ['viewClassMap' => ['xlsx' => 'CakeSpreadsheet.Spreadsheet'], 'enableBeforeRedirect' => false]);
```

Each application *must* have an xlsx layout. The following is a barebones layout that can be placed in `src/Template/Layout/xlsx/default.ctp`:

```php
<?= $this->fetch('content') ?>
```

Finally, you can link to the current page with the .xlsx extension. This assumes you've created an `xlsx/index.ctp` file in your particular controller's template directory:

```php
$this->Html->link('Excel file', ['_ext' => 'xlsx']);
```

Inside your view file you will have access to the PHPSpreadsheet library with `$this->Spreadsheet`. Please see the [PHPSpreadsheet](https://github.com/PHPOffice/PHPSpreadsheet) documentation for a guide on how to use PHPSpreadsheet.

## License

The MIT License (MIT)

Copyright (c) Christakis Mina

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.
