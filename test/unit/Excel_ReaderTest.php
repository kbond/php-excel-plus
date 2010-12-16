<?php

require_once dirname(__FILE__).'/../../lib/Excel_Reader.php';

/**
 * Test class for Excel_Reader.
 * Generated by PHPUnit on 2010-12-14 at 09:28:52.
 */
class Excel_ReaderTest extends PHPUnit_Framework_TestCase
{
  /**
   * @var Excel_Reader
   */
  protected $object;

  protected $testfile;

  /**
   * Sets up the fixture, for example, opens a network connection.
   * This method is called before a test is executed.
   */
  protected function setUp()
  {
    $this->object = new Excel_Reader;
    $this->testfile = dirname(__FILE__).'/../fixtures/simple.xls';
  }
  
  public function testLoad()
  {
    $this->setExpectedException('Exception');
    $this->object->load('non existant file');
  }

  public function testGetPHPExcel()
  {
    $this->object->load($this->testfile);

    $this->assertType('PHPExcel', $this->object->getPHPExcel(), 'Type correct');
  }

  public function testSimpleHydrate()
  {
  }
}
