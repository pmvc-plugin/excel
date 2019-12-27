<?php
namespace PMVC\PlugIn\excel;

${_INIT_CONFIG}[_CLASS] = __NAMESPACE__.'\excel';

class excel extends \PMVC\PlugIn
{
    public function init()
    {
    }

    public function create()
    {
        \PMVC\l(__DIR__.'/src/XLSXWriter.php');
        $writer = new \XLSXWriter();
        return $writer;
    }

    public function read($f)
    {
        \PMVC\l(__DIR__.'/src/excel_reader2.php');
        \PMVC\l(__DIR__.'/src/SpreadsheetReader.php');
        $reader = new \SpreadsheetReader($f);
        return $reader;
    }
}
