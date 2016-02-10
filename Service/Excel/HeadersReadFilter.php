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
 * Use HeadersReadFilter to read only the first row of an Excel worksheet.
 **
 * @author Christian Mattart <christian@chmat.be>
 */
class HeadersReadFilter implements \PHPExcel_Reader_IReadFilter
{
    /**
     * Return true if specified cell should be read.
     * 
     * This will read row 1 (usually column headings)
     * 
     * @param integer $column
     * @param integer $row
     * @param string $worksheetName
     * @return boolean
     */
    public function readCell($column, $row, $worksheetName = '')
    {
        //  Only read the first row
        return ($row == 1);
    }

}
