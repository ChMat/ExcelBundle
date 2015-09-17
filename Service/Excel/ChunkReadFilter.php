<?php

/*
 * This file is part of ChMatExcelBundle.
 *
 * (c) Christian Mattart <christian@chmat.be>
 *
 * This source file is subject to the MIT license that is bundled
 * with this source code in the LICENSE file.
 */

namespace ChMat\ExcelBundle\Service\Excel;

/**
 * Use a ChunkReadFilter to read an Excel worksheet in small chunks.
 * 
 * This filter is used to read a specific number of rows.
 *
 * @author Christian Mattart <christian@chmat.be>
 */
class ChunkReadFilter implements \PHPExcel_Reader_IReadFilter
{

    private $_startRow = 0;
    private $_endRow = 0;

	/**
	 * Set the row range that we want to read.
	 * 
	 * @param int $startRow
	 * @param int $chunkSize
	 */
    public function __construct($startRow, $chunkSize)
    {
        $this->setRows($startRow, $chunkSize);
    }

	/**
	 * Set the row range that we want to read.
	 * 
	 * @param int $startRow
	 * @param int $chunkSize Number of rows to read from $startRow
	 */
    public function setRows($startRow, $chunkSize)
    {
        $this->_startRow = $startRow;
        $this->_endRow = $startRow + $chunkSize;
    }

    /**
     * Return true if specified cell should be read.
     * 
     * This will read row 1 (usually column headings)
     * and rows specified in $this->setRows().
     * 
     * @param integer $column
     * @param integer $row
     * @param string $worksheetName
     * @return boolean
     */
    public function readCell($column, $row, $worksheetName = '')
    {
        //  Only read the heading row, and the rows that are configured in $this->_startRow and $this->_endRow
        if ($row == 1 || ($row >= $this->_startRow && $row <= $this->_endRow))
        {
            return true;
        }
        return false;
    }

}
