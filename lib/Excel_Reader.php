<?php

require_once dirname(__FILE__).'/vendor/PHPExcel/PHPExcel.php';

class Excel_Reader
{
  /** @var PHPExcel */
  protected $objPHPExcel;

  /**
   * Constructor
   *
   * @param string $filename filename to load (optional)
   */
  public function __construct($filename = null)
  {
    if ($filename)
      $this->load($filename);
  }

  /**
   * Loads in excel file
   *
   * @param string $filename filename to load
   * @return Excel_Reader
   */
  public function load($filename)
  {
    $this->objPHPExcel = PHPExcel_IOFactory::load($filename);

    return $this;
  }

  /**
   * Return instance of PHPExcel
   *
   * @return PHPExcel
   */
  public function getPHPExcel()
  {
    return $this->objPHPExcel;
  }
}