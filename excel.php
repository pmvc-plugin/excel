<?php
namespace PMVC\PlugIn\excel;

// \PMVC\l(__DIR__.'/xxx.php');

${_INIT_CONFIG}[_CLASS] = __NAMESPACE__.'\excel';

class excel extends \PMVC\PlugIn
{
    public function init()
    {
        \PMVC\l(__DIR__.'/src/XLSXWriter.php');
    }

    public function create()
    {
        $writer = new \XLSXWriter();
        return $writer;
    }
}
