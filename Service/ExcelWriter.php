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

use ChMat\ExcelBundle\Exception\CannotReadFileException;
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
    protected $objPHPExcel;


    /**
     * Filename to read.
     * 
     * @var string
     */
    protected $filename;

    /**
     * Type of file to read.
     * 
     * Value is one of the constants in \PHPExcel_Reader_IReader.
     *
     * @var string
     */
    protected $fileType;

    /**
     * Indique si un fichier est chargé dans le lecteur.
     *
     * @var boolean
     */
    protected $isLoaded;

    /**
     * 1-based row index.
     * 
     * This value is not reset by the object when switching between sheets.
     *
     * @var integer
     */
    protected $rowIndex;

    /**
     * Nombre de colonnes dans la feuille active.
     *
     * @var integer
     */
    protected $numberOfCols;

    /**
     * Nombre de lignes dans la feuille active.
     *
     * @var integer
     */
    protected $numberOfRows;

    /**
     *
     * @var \PHPExcel_Reader_IReader
     */
    protected $reader;

    /**
     * @var \PHPExcel_Writer_IWriter
     */
    protected $writer;

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

        $this->rowIndex = 1;
        $this->isLoaded = true;

        return $this;
    }

    /**
     * Load an existing file in memory in order to write into it and instantiate PhpExcel object.
     * 
     * @param string $filename Chemin du fichier à charger
     * @param null|string $fileType Type de fichier (ou null si la confiance règne).
     *
     * @return $this
     * @throws CannotReadFileException
     * @throws \PHPExcel_Reader_Exception
     */
    public function loadFile($filename, $fileType = null)
    {
        if (null !== $fileType)
        {
            $this->fileType = $fileType;
        }
        else
        {
            $this->fileType = \PHPExcel_IOFactory::identify($filename);
        }

        if (!is_readable($filename))
        {
            throw new CannotReadFileException($filename);
        }

        $this->filename = $filename;

        $this->reader = \PHPExcel_IOFactory::createReader($this->fileType);

        if ($this->fileType !== 'CSV')
        {
            $this->reader->setReadDataOnly(false);
        }
        else
        {
            $this->reader->setDelimiter(';');
        }

        $this->objPHPExcel = $this->reader->load($this->filename);

        $this->countRows();
        $this->rowIndex = 1;
        $this->isLoaded = true;

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
        $this->rowIndex = 1;

        if ($useKeysAsColumnNames)
        {
            // Setting titles
            $columnIndex = 0;
            reset($data);

            foreach ($data[key($data)] as $columnName => $row)
            {
                if (is_object($row))
                    continue;
                $this->setCellValueInActiveSheet($columnIndex, $this->rowIndex, $columnName);
                $this->objPHPExcel->getActiveSheet()->getColumnDimensionByColumn($columnIndex)->setAutoSize(true);
                $columnIndex++;
            }

            reset($data);
            $this->rowIndex++;
        }
        
        $this->populateRows($data);
        
        return $this;
    }

    /**
     * Append $data to an existing sheet.
     *
     * @param array $data
     *
     * @return ExcelWriter
     * @throws \Exception
     */
    public function appendData($data)
    {
        if (!$this->isLoaded)
        {
            throw new \Exception('Excel file was not loaded.');
        }

        $this->rowIndex = $this->numberOfRows + 1;

        $this->populateRows($data);

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
     * By default, it will not overwrite an existing file with the same name.
     *
     * @param string $filename Filename with path
     * @param string $format One of xls|xlsx
     * @param bool $canOverWrite Set to true if you want to let an existing file be overwritten upon saving.
     *
     * @return bool
     * @throws FileExistsException
     * @throws \PHPExcel_Reader_Exception
     */
    public function saveFile($filename, $format, $canOverWrite = false)
    {
        switch ($format)
        {
            case 'xlsx':
                $format = 'Excel2007';
                break;
            case 'xls':
                $format = 'Excel5';
                break;
        }

        $objWriter = \PHPExcel_IOFactory::createWriter($this->objPHPExcel, $format);

        if (!$canOverWrite and file_exists($filename))
        {
            throw new FileExistsException();
        }

        $objWriter->save($filename);

        $this->objPHPExcel->disconnectWorksheets();

        unset($this->objPHPExcel);

        return true;
    }

    /**
     * Returns the number of columns in current sheet.
     *
     * @return int
     */
    public function getNumberOfCols()
    {
        if (!$this->isLoaded) return false;

        if (null === $this->numberOfCols) $this->countColumns();

        return $this->numberOfCols;
    }

    /**
     * Retourne le nombre de lignes dans la feuille active.
     *
     * @return int
     */
    public function getNumberOfRows()
    {
        if (!$this->isLoaded) return false;

        if (null === $this->numberOfRows) $this->countRows();

        return $this->numberOfRows;
    }

    /**
     * Retourne l'indice de la ligne courante
     *
     * @return int Numéro de la ligne courante, à partir de 1
     */
    public function getRowIndex()
    {
        return $this->rowIndex;
    }

    /**
     * Change la ligne courante
     *
     * @param integer $rowIndex
     * @return ImportReader
     */
    public function setRowIndex($rowIndex)
    {
        $this->rowIndex = $rowIndex;

        return $this;
    }

    ////////////////////////////////////////////////////////////////////////////
    // Protected Methods
    ////////////////////////////////////////////////////////////////////////////
    
    /**
     * Compte le nombre de lignes contenues dans le fichier.
     */
    protected function countRows()
    {
        $excelReader = $this->reader->load($this->filename);
        $this->numberOfRows = $excelReader->getActiveSheet()->getHighestRow();

        $excelReader->disconnectWorksheets();
        unset($excelReader);
    }

    /**
     * Compte le nombre de colonnes contenues dans le fichier.
     */
    protected function countColumns()
    {
        $excelReader = $this->reader->load($this->filename);
        $this->numberOfCols = $excelReader->getActiveSheet()->getHighestColumn();

        $excelReader->disconnectWorksheets();
        unset($excelReader);
    }

    /**
     * Inject $data values into sheet starting at current $this->rowIndex.
     * 
     * @param array $data
     */
    protected function populateRows($data)
    {
        foreach ($data as $rowData)
        {
            $columnIndex = 0;
            foreach ($rowData as $cellValue)
            {
                if (is_object($cellValue))
                {
                    continue;
                }

                $this->setCellValueInActiveSheet($columnIndex, $this->rowIndex, $cellValue);

                $columnIndex++;
            }

            $this->rowIndex++;
        }
    }

    /**
     * Set value $value in cell with coordinates $cellRow, $cellColumn in active sheet.
     *
     * @param int $cellColumn Column index (0-indexed)
     * @param int $cellRow Row index (1-indexed)
     * @param mixed $value Value to be used (cannot be an object)
     */
    protected function setCellValueInActiveSheet($cellColumn, $cellRow, $value)
    {
        // A bit of DIY here in order to make it such as that 0 is considered a numeric value.
        if ((is_numeric($value) && (substr($value, 0, 1) != 0)) || (is_numeric($value) && strlen($value) == 1))
        { // A number not-starting with 0 or a single-digit number
            $this->objPHPExcel->getActiveSheet()->setCellValueExplicitByColumnAndRow($cellColumn,
                $cellRow, $value,
                \PHPExcel_Cell_DataType::TYPE_NUMERIC);
        }
        elseif (is_numeric($value) && (substr($value, 0, 1) == 0))
        { // A number starting with 0
            $this->objPHPExcel->getActiveSheet()->setCellValueExplicitByColumnAndRow($cellColumn,
                $cellRow, $value,
                \PHPExcel_Cell_DataType::TYPE_STRING);
        }
        else
        { // Any other value will be considered as a string
            $this->objPHPExcel->getActiveSheet()->setCellValueExplicitByColumnAndRow($cellColumn,
                $cellRow, $value,
                \PHPExcel_Cell_DataType::TYPE_STRING);
        }
    }

}
