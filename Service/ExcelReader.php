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

use ChMat\ExcelBundle\Service\Excel\ChunkReadFilter;
use ChMat\ExcelBundle\Exception\CannotReadFileException;

/**
 * Simple Excel Worksheet reader.
 * 
 * This is used to extract the contents of an Excel worksheet into a two-dimensional array.
 *
 * @author Christian Mattart <christian@chmat.be>
 */
class ExcelReader
{
    /*
     * Définition des types de fichiers reconnus
     */
    const EXCEL5 = 'Excel5';
    const EXCEL2007 = 'Excel2007';
    const EXCEL2003 = 'Excel2003XML';
    const OOCALC    = 'OOCalc';
    const GNUMERIC  = 'Gnumeric';

    /**
     * Nombre de lignes lues par appel à readNextRows(). 20 par défaut.
     * 
     * @var integer
     */
    private $chunkSize;

    /**
     * Nom du fichier à lire
     * @var string
     */
    private $filename;
    
    /**
     * Type du fichier à lire (parmi les constantes).
     * 
     * @var string
     */
    public $fileType;
    
    /**
     * Indique si un fichier est chargé dans le lecteur.
     * 
     * @var boolean
     */
    private $isLoaded;
    
    
    /**
     * 1-based row index.
     * 
     * @var integer
     */
    private $rowIndex;
    
    /**
     * Nombre de colonnes dans la feuille active.
     * 
     * @var integer
     */
    private $numberOfCols;
    
    /**
     * Nombre de lignes dans la feuille active.
     * 
     * @var integer
     */
    private $numberOfRows;
    
    /**
     *
     * @var type 
     */
    private $reader;

    /**
     *
     * @var ChunkReadFilter
     */
    private $chunkFilter;

    

    
    ////////////////////////////////////////////////////////////////////////////
    // Public Methods
    ////////////////////////////////////////////////////////////////////////////

    
    

    public function __construct()
    {
        set_time_limit(0);
        
        $this->reset();
    }
    
    /**
     * Returns an array with the contents of current sheet first row.
     * 
     * @return array
     */
    public function getColumnHeaders()
    {
        if (!$this->isLoaded) return false;
        
        $this->applyFilter(1, 1);

        $result = $this->reader->load($this->filename);

        $headers = $result->getActiveSheet()->toArray(null, true, true, true);
        
        return $headers[1];
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
     * Retourne l'index de la ligne courante.
     * 
     * @return integer
     */
    public function getRowIndex()
    {
        return $this->rowIndex;
    }


    /**
     * Charge un fichier $filename et active l'objet PHPExcel en mesure de le lire.
     * Si $fileType n'est pas défini, le type de fichier sera déterminé automatiquement.
     * 
     * @param string $filename
     * @param string $fileType
     * @throws CannotReadFileException
     * @return $this
     */
    public function load($filename, $fileType = null)
    {
        //ini_set('memory_limit', '256M');
        
        $this->reset();
        
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
            $this->reader->setReadDataOnly(true);
        }
        else
        {
            $this->reader->setDelimiter(';');
        }
        
        $this->rowIndex = 1;
        
        //$this->countRows();
        
        $this->isLoaded = true;
        
        return $this;
    }

    /**
     * Retourne un morceau du fichier ($chunkSize lignes à partir de la ligne $startAtRow).
     * 
     * Si les paramètres $chunkSize et $startAtRow ne sont pas définis, 20 lignes sont retournées
     * à chaque appel de la fonction.
     * 
     * @param integer $chunkSize
     * @param integer $startAtRow
     * @return array
     */
    public function readNextRows($chunkSize = null, $startAtRow = null)
    {
        if (null !== $startAtRow)
        {
            $this->rowIndex = $startAtRow;
        }
        
        if (null !== $chunkSize)
        {
            $this->chunkSize = $chunkSize;
        }
        $currentChunkSize = ($this->rowIndex + $this->chunkSize > $this->numberOfRows) ? $this->numberOfRows - $this->rowIndex + 1 : $this->chunkSize;

        // Si on est au bout du fichier ou au-delà, on s'arrête.
        if ($currentChunkSize <= 0)
        {
            return false;
        }

        $this->applyFilter($this->rowIndex, $currentChunkSize);
        
        /*
         * Load only the rows that match our filter from $inputFileName to a PHPExcel Object  
         */
        $result = $this->reader->load($this->filename);
        
        // On retourne les données.
        $temp = $result->getActiveSheet()->toArray(null, true, true, true);
        $return = array();
        
        /*
         * For some bizarre reason, even if readFilter is applied, returned results still contain
         * all rows that should not be read. They are empty, but nonetheless.
         * 
         * TODO Check if this is normal behaviour.
         */
//        $f = fopen('app/logs/extraction.txt', 'a');
//        fwrite($f, sprintf('========BEGIN===== [rowIndex : %d - chunkSize : %d - count : %d ]=========================', $this->rowIndex, $this->chunkSize, sizeof($temp)));
//        fwrite($f, print_r($temp, true));
//        fwrite($f, '========END================================');
//        fclose($f);
        $outputIndex = 1;
        for ($row = $this->rowIndex; $row < $this->rowIndex + $currentChunkSize; $row++)
        {
            $return[$outputIndex] = $temp[$row];
            $outputIndex++;
        }
        
        unset($temp);
        
        // On prépare l'index de ligne pour la lecture suivante.
        $this->rowIndex += $this->chunkSize;

        $result->disconnectWorksheets();
        unset($result);
        
        return $return;
    }
    
    /**
     * Retourne le contenu de toute la feuille active ou uniquement les lignes $fromRow à $toRow.
     * 
     * fixme les items sont groupés par groupes de 20 lignes (tableau à 3 dimensions)
     * fixme 
     * 
     * @param integer $fromRow
     * @param integer $toRow
     * @return array
     */
    public function read($fromRow = null, $toRow = null)
    {
        $previousRowIndex = $this->rowIndex;
        
        $this->rowIndex = $fromRow = ($fromRow !== null) ? $fromRow : 1;
        
        if ($toRow === null)
        {    
            $this->countRows();
            $toRow = $this->numberOfRows;
        }
        
        $sheetData = array();
        
        for ($this->rowIndex = $fromRow; $this->rowIndex <= $toRow; $this->rowIndex += $this->chunkSize)
        {
            
            $this->applyFilter($this->rowIndex, $this->chunkSize);

            $result = $this->reader->load($this->filename);

            $chunk = $result->getActiveSheet()->toArray(null, true, true, true);
	        
	        foreach ($chunk as $row => $data)
	        {
		        $sheetData[$row] = $data;
	        }
        }
        
        $this->rowIndex = $previousRowIndex;
        
        $result->disconnectWorksheets();
        unset($result);
                
        return $sheetData;
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
     * Applique le filtre au lecteur.
     * @param integer $fromRow
     * @param integer $chunkSize
     */
    protected function applyFilter($fromRow, $chunkSize)
    {
        if (!$this->chunkFilter instanceof ChunkReadFilter)
        {
            $this->chunkFilter = new ChunkReadFilter($fromRow, $chunkSize);
            $this->reader->setReadFilter($this->chunkFilter);
        }
        else
        {
            $this->chunkFilter->setRows($fromRow, $chunkSize);
        }

        
    }
    
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
     * Réinitialise le lecteur à la construction et au chargement d'un fichier.
     */
    protected function reset()
    {
        $this->chunkFilter = null;
        $this->chunkSize = 20;
        $this->fileType = null;
        $this->filename = null;
        $this->isLoaded = false;
        $this->numberOfRows = null;
        $this->rowIndex = 1;
    }
}
