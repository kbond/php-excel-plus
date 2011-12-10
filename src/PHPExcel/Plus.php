<?php

/**
 * Wrapper for PHPExcel
 *
 * + Loads in an excel file
 * + Perform array hydrations
 *
 * @package PHPExcel
 * @author kbond
 */
class PHPExcel_Plus
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
        if ($filename) {
            $this->load($filename);
        }
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
    public function convertToSimpleArray($sheet = 0)
    {
        $sheet = $this->getSheet($sheet);

        $rows = $sheet->toArray();
        $return = array();

        // grab headers
        $headers = array_shift($rows);

        foreach ($rows as $row_num => $row) {
            if (trim($row[0]) == null) {
                continue;
            }

            foreach ($row as $key => $value) {
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

    /**
     * Returns a complex array with the following features:
     *
     * + Cell rows/cols count (merged cells)
     * + Frozen panes split between table header/body
     * + Bold text cells
     *
     * This is useful for maping an excel table to an html table
     *
     * @param string $sheet the name of the sheet
     * @return array
     */
    public function convertToComplexArray($sheet = 0)
    {
        $sheet = $this->getSheet($sheet);

        $return = array();

        // get frozen panes if set
        $header = $sheet->getFreezePane();

        /* @var $row PHPExcel_Worksheet_Row */
        foreach ($sheet->getRowIterator() as $row) {
            $table_section = 'body';

            // see if row is in header (above frozen pane)
            if ($header) {
                $header_end = PHPExcel_Cell::coordinateFromString($sheet->getFreezePane());
                $header_end = $header_end[1];

                if ($row->getRowIndex() < $header_end) {
                    // row is in table header
                    $table_section = 'head';
                }
            }

            /* @var $cell PHPExcel_Cell */
            foreach ($row->getCellIterator() as $cell) {
                $info = array();
                // check if cell merged - add rows/cols
                if ($range = $this->getMergedRange($sheet, $cell)) {
                    // split range
                    $range_details = PHPExcel_Cell::splitRange($range);

                    // check if not first cell in merged ranged (top-left)
                    if ($range_details[0][0] != $cell->getCoordinate()) {
                        continue;
                    }

                    // get range dimension
                    $range_dimension = PHPExcel_Cell::rangeDimension($range);

                    // set cols
                    if ($range_dimension[0] > 1) {
                        $info['cols'] = $range_dimension[0];
                    }

                    // set rows
                    if ($range_dimension[1] > 1) {
                        $info['rows'] = $range_dimension[1];
                    }
                }

                // set cell value
                $info['value'] = $cell->getValue();

                // additional info
                $info = $this->setCellInfo($info, $sheet, $row, $cell);

                $return[$table_section][$row->getRowIndex()][$cell->getColumn()] = $info;
            }
        }

        return $return;
    }

    /**
     * Add additional cell info (override to add your own)
     *
     * + bold flag
     * + italic flag
     * + color hex
     * + bgcolor hex
     * + alignment
     *
     * @param array $info
     * @param PHPExcel_Worksheet $sheet
     * @param PHPExcel_Worksheet_Row $row
     * @param PHPExcel_Cell $cell
     */
    public function setCellInfo($info, PHPExcel_Worksheet $sheet, PHPExcel_Worksheet_Row $row, PHPExcel_Cell $cell)
    {
        /* @var $style PHPExcel_Style */
        $style = $sheet->getStyle($cell->getCoordinate());

        /* @var $font PHPExcel_Style_Font */
        $font = $style->getFont();

        /* @var $fill PHPExcel_Style_Fill */
        $fill = $style->getFill();

        /* @var $alignment PHPExcel_Style_Alignment */
        $alignment = $style->getAlignment();

        // set bold flag
        $info['bold'] = $font->getBold();

        // set italic flag
        $info['italic'] = $font->getItalic();

        // set underline flag
        $info['underline'] = $font->getUnderline();

        // set alignment
        $info['align'] = $alignment->getHorizontal();

        // set background color
        $info['bgcolor'] = $fill->getStartColor()->getRGB();

        // set font color
        $info['color'] = $font->getColor()->getRGB();

        return $info;
    }

    /**
     * Check the sheet to see if the specified cell is in it's merged cell list
     *
     * @param PHPExcel_Worksheet $sheet
     * @param PHPExcel_Cell $cell
     * @return the range if found or false
     */
    protected function getMergedRange(PHPExcel_Worksheet $sheet, PHPExcel_Cell $cell)
    {
        foreach ($sheet->getMergeCells() as $range) {
            if ($cell->isInRange($range)) {

                return $range;
            }
        }

        return false;
    }

    /**
     * Returns the worksheet by name or number
     *
     * @param string|int $sheet
     * @return PHPExcel_Worksheet
     */
    protected function getSheet($sheet = 0)
    {
        if (is_string($sheet)) {
            $sheet = $this->objPHPExcel->getSheetByName($sheet);
        } else {
            // get first sheet if none defined
            $sheet = $this->objPHPExcel->getSheet($sheet);
        }

        return $sheet;
    }

}