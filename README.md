[![Latest Stable Version](https://poser.pugx.org/pmvc-plugin/excel/v/stable)](https://packagist.org/packages/pmvc-plugin/excel) 
[![Latest Unstable Version](https://poser.pugx.org/pmvc-plugin/excel/v/unstable)](https://packagist.org/packages/pmvc-plugin/excel) 
[![Build Status](https://travis-ci.org/pmvc-plugin/excel.svg?branch=master)](https://travis-ci.org/pmvc-plugin/excel)
[![License](https://poser.pugx.org/pmvc-plugin/excel/license)](https://packagist.org/packages/pmvc-plugin/excel)
[![Total Downloads](https://poser.pugx.org/pmvc-plugin/excel/downloads)](https://packagist.org/packages/pmvc-plugin/excel) 

Excel Library (for now write only)
===============
   * Fork from https://github.com/mk-j/PHP_XLSXWriter
   * Replace with pclzip don't need compile ZipArchive in this version
      * http://php.net/manual/en/zip.installation.php 

## Office Open XML File Formats
   * https://msdn.microsoft.com/en-us/library/aa338205(v=office.12).aspx
   * https://en.wikipedia.org/wiki/Office_Open_XML_file_formats

## How to use?
```
include_once('vendor/autoload.php');
PMVC\Load::plug();

$data = array(
    array('year','month','amount'),
    array('2003','1','220'),
    array('2003','2','153.5'),
);

$writer = \PMVC\plug('excel')->create();
$writer->writeSheet($data);
$writer->writeToFile('output.xlsx');
```


## Install with Composer
### 1. Download composer
   * mkdir test_folder
   * curl -sS https://getcomposer.org/installer | php

### 2. Install by composer.json or use command-line directly
#### 2.1 Install by composer.json
   * vim composer.json
```
{
    "require": {
        "pmvc-plugin/excel": "dev-master"
    }
}
```
   * php composer.phar install

#### 2.2 Or use composer command-line
   * php composer.phar require pmvc-plugin/excel


