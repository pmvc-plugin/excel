[![Latest Stable Version](https://poser.pugx.org/pmvc-plugin/excel/v/stable)](https://packagist.org/packages/pmvc-plugin/excel) 
[![Latest Unstable Version](https://poser.pugx.org/pmvc-plugin/excel/v/unstable)](https://packagist.org/packages/pmvc-plugin/excel) 
[![Build Status](https://travis-ci.org/pmvc-plugin/excel.svg?branch=master)](https://travis-ci.org/pmvc-plugin/excel)
[![License](https://poser.pugx.org/pmvc-plugin/excel/license)](https://packagist.org/packages/pmvc-plugin/excel)
[![Total Downloads](https://poser.pugx.org/pmvc-plugin/excel/downloads)](https://packagist.org/packages/pmvc-plugin/excel) 

Excel Library (for now write only)
===============

Fork from https://github.com/mk-j/PHP_XLSXWriter

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
        "pmvc-plugin/hello_world": "dev-master"
    }
}
```
   * php composer.phar install

#### 2.2 Or use composer command-line
   * php composer.phar require pmvc-plugin/hello_world

### 3. Write some demo code
```
<?php
    include_once('vendor/pmvc/pmvc/include_plug.php');
    PMVC\setPlugInFolder('vendor/pmvc-plugin/');
    PMVC\plug('hello_world')->say('hello, World!');
?>
```
### 4. Run the demo
   * php demo.php

### 5. Check the whole demo code
   * https://github.com/pmvc-plugin/hello_world/tree/master/demo

