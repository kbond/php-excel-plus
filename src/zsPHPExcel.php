<?php

require_once dirname(__FILE__) . '/../vendor/PHPExcel/PHPExcel.php';

/**
 * Wrapper for PHPExcel
 *
 * + Loads in an excel file
 * + Perform array hydrations
 *
 * @package php-excel
 * @author kbond
 */
class zsPHPExcel
{

  /**
   * @var PHPExcel 
   */
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

  /**
   * Converts the specified worksheet into a column header keyed array.
   *
   * Useful for importing excel files into database.
   *
   * If no sheet is specified the first one is used. Skips row if first column
   * blank.
   *
   * Example return:
   * array(
   *    0 => array(
   *      'A1' => 'A2',
   *      'B1' => 'B2'),
   *    1 => array(
   *      'A1' => 'A3',
   *      'B1' => 'B3'),
   *    2 => array(
   *      'A1' => 'A4',
   *      'B1' => 'B4')
   *  );
   *
   * @param string $sheet the name of the sheet
   * @return array
   */
  public function convertToSimpleArray($sheet = null)
  {
    // get first sheet if none defined
    if (!$sheet)
      $sheet = $this->objPHPExcel->getSheet();
    else
      $sheet = $this->objPHPExcel->getSheetByName($sheet);

    $rows = $sheet->toArray();
    $return = array();

    // grab headers
    $headers = array_shift($rows);

    foreach ($rows as $row_num => $row)
    {
      if (trim($row[0]) == null)
        continue;

      foreach ($row as $key => $value)
      {
        try {
          $return[$row_num][$headers[$key]] = trim($value);
        } catch (Exception $e) {
          //Number of column headers do not match
          continue;
        }
      }
    }

    return $return;
  }

  public function convertToComplexArray($sheet = null)
  {
    // get first sheet if none defined
    if (!$sheet)
      $sheet = $this->objPHPExcel->getSheet();
    else
      $sheet = $this->objPHPExcel->getSheetByName($sheet);

    
    foreach ($sheet->getRowIterator() as $row)
    {
      /* @var $cell PHPExcel_Cell */
      foreach ($row->getCellIterator() as $cell)
      {
        die(var_dump($cell->getRangeBoundaries()));
      }
    }
    
  }

}