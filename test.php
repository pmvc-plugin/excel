<?php
PMVC\Load::plug();
PMVC\addPlugInFolder('../');
class ExcelTest extends PHPUnit_Framework_TestCase
{
    private $_plug = 'excel';
    function testPlugin()
    {
        ob_start();
        print_r(PMVC\plug($this->_plug));
        $output = ob_get_contents();
        ob_end_clean();
        $this->assertContains($this->_plug,$output);
    }

    function testExcel()
    {
        $excel = PMVC\plug($this->_plug);
        $data = array(
                array('year','month','amount'),
                array('2003','1','220'),
                array('2003','2','153.5'),
        );

        $writer = $excel->create();
        $writer->writeSheet($data);
        $writer->writeToFile('output.xlsx');
    } 

}
