<?php

/*
 * This file is part of ChMatExcelBundle.
 *
 * (c) Christian Mattart <christian@chmat.be>
 *
 * This source file is subject to the MIT license that is bundled
 * with this source code in the LICENSE file.
 */

namespace ChMat\ExcelBundle\Service;

use ChMat\ExcelBundle\Exception\FileExistsException;

/**
 * Simple Excel Worksheet Writer.
 * 
 * This is used to save a two-dimensional array into an Excel worksheet.
 * 
 * @author Jean-François de Locht
 * @author Christian Mattart <christian@chmat.be>
 */
class ExcelWriter
{

    /**
     * PHPExcel Object
     * 
     * @var \PHPExcel
     */
    private $objPHPExcel;

    /**
     * Create $this->objPHPExcel
     * Set document name, title, author and locale.
     * 
     * @param string $name Document name (not used as filename)
     * @param string $title Document title
     * @param string $author Author name
     * @param string $locale Document locale (default fr_FR)
     * @return ExcelWriter
     * @throws \Exception
     */
    public function createDocument($name = '', $title = '', $author = '',
            $locale = 'fr_FR')
    {
        $this->objPHPExcel = new \PHPExcel();

        //         $cacheMethod = \PHPExcel_CachedObjectStorageFactory::cache_to_apc;
        //         $cacheSettings = array( 'cacheTime' => 600);
        //         \PHPExcel_Settings::setCacheStorageMethod($cacheMethod, $cacheSettings);
        //         $cacheMethod = \PHPExcel_CachedObjectStorageFactory::cache_to_discISAM;
        //         \PHPExcel_Settings::setCacheStorageMethod($cacheMethod);
        $cacheMethod = \PHPExcel_CachedObjectStorageFactory:: cache_to_phpTemp;
        $cacheSettings = array(' memoryCacheSize ' => '1024MB');
        if (!\PHPExcel_Settings::setCacheStorageMethod($cacheMethod,
                        $cacheSettings))
        {
            throw new \Exception('Unable to change cache storage method. Please contact your administrator.');
        }

        $this->objPHPExcel->getProperties()->setCreated($author);
        $this->objPHPExcel->getProperties()->setLastModifiedBy($author);
        $this->objPHPExcel->getProperties()->setTitle($title);
        $this->objPHPExcel->getProperties()->setCreator($author);

        $validLocale = \PHPExcel_Settings::setLocale($locale);

        $this->gotoSheetByIndex(0);
        
        return $this;
    }

    /**
     * Jump to sheet with numeric $index.
     * 
     * @param int $index
     * @return ExcelWriter
     */
    public function gotoSheetByIndex($index = 0)
    {
        $this->objPHPExcel->setActiveSheetIndex(0);

        return $this;
    }

    /**
     * Jump to sheet with string $title.
     * 
     * @param string $title
     * @return ExcelWriter
     */
    public function gotoSheetByTitle($title)
    {
        $this->objPHPExcel->setActiveSheetIndexByName($title);

        return $this;
    }

    /**
     * Change current sheet $title.
     * 
     * @param string $title
     * @return ExcelWriter
     */
    public function setSheetTitle($title)
    {
        $this->objPHPExcel->getActiveSheet()->setTitle($title);

        return $this;
    }

    /**
     * Fill current sheet with $data.
     * 
     * @param array $data
     * @param boolean $useKeysAsColumnNames If true, first row will be populated by index value of each column.
     * @return ExcelWriter
     */
    public function fillSheet($data, $useKeysAsColumnNames = true)
    {
        $cellRow = 1;

        if ($useKeysAsColumnNames)
        {
            // Setting titles
            $i = 0;
            reset($data);

            foreach ($data[key($data)] as $key => $row)
            {
                if (is_object($row))
                    continue;
                $this->objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow($i,
                        $cellRow, $key);
                $this->objPHPExcel->getActiveSheet()->getColumnDimensionByColumn($i)->setAutoSize(true);
                $i++;
            }

            reset($data);
            $cellRow++;
        }

        foreach ($data as $key => $value)
        {

            $cellCol = 0;
            foreach ($value as $colName => $colValue)
            {
                if (is_object($colValue))
                    continue;
                // A bit of DIY here in order to make it such as that 0 is considered a numeric value.
                if ((is_numeric($colValue) && (substr($colValue, 0, 1) != 0)) || (is_numeric($colValue) && strlen($colValue) == 1))
                {
                    $this->objPHPExcel->getActiveSheet()->setCellValueExplicitByColumnAndRow($cellCol,
                            $cellRow, $colValue,
                            \PHPExcel_Cell_DataType::TYPE_NUMERIC);
                }
                elseif (is_numeric($colValue) && (substr($colValue, 0, 1) == 0))
                {
                    $this->objPHPExcel->getActiveSheet()->setCellValueExplicitByColumnAndRow($cellCol,
                            $cellRow, $colValue,
                            \PHPExcel_Cell_DataType::TYPE_STRING);
                }
                else
                {
                    $this->objPHPExcel->getActiveSheet()->setCellValueExplicitByColumnAndRow($cellCol,
                            $cellRow, $colValue,
                            \PHPExcel_Cell_DataType::TYPE_STRING);
                }
                $cellCol++;
            }

            $cellRow++;
        }

        return $this;
    }

    /**
     * Send file to client.
     * 
     * @param string $filename Filename without extension
     * @param string $format One of xls|xlsx
     */
    public function outputFile($filename, $format)
    {
        $filename = sprintf('%s.%s', str_replace(' ', '_', $filename), $format);

        if ($format == 'xlsx')
        {
            header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            header('Content-Disposition: attachment;filename="' . $filename. '"');
            header('Cache-Control: max-age=0');
            $objWriter = new \PHPExcel_Writer_Excel2007($this->objPHPExcel);
        }
        else
        {
            header('Content-Type: application/vnd.ms-excel');
            header('Content-Disposition: attachment;filename="' . $filename . '"');
            header('Cache-Control: max-age=0');
            $objWriter = \PHPExcel_IOFactory::createWriter($this->objPHPExcel,
                            'Excel5');
        }
        $objWriter->save('php://output');
        $this->objPHPExcel->disconnectWorksheets();

        unset($this->objPHPExcel);

        exit;
    }

	/**
	 * Save file to disk.
	 *
	 * @param string $filename Filename with path
	 * @param string $format One of xls|xlsx
	 *
	 * @return bool
	 * @throws FileExistsException
	 * @throws \PHPExcel_Reader_Exception
	 */
    public function saveFile($filename, $format)
    {
        if ($format == 'xlsx')
        {
            $objWriter = new \PHPExcel_Writer_Excel2007($this->objPHPExcel);
        }
        else
        {
            $objWriter = \PHPExcel_IOFactory::createWriter($this->objPHPExcel, 'Excel5');
        }
	    
	    if (file_exists($filename))
	    {
		    throw new FileExistsException();
	    }
	    		    
        $objWriter->save($filename);
	    
        $this->objPHPExcel->disconnectWorksheets();

        unset($this->objPHPExcel);
	    
	    return true;
    }

}
