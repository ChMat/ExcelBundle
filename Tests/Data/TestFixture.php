<?php

/*
 * This file is part of ChMatExcelBundle.
 *
 * (c) Christian Mattart <christian@chmat.be>
 *
 * This source file is subject to the MIT license that is bundled
 * with this source code in the LICENSE file.
 */

namespace ChMat\ExcelBundle\Tests\Data;

/**
 * Class TestFixture
 * 
 * This fixture contains the data that is stored in the `Tests/Data/test_read.xls` file.
 * It is used to test this bundle.
 * 
 * @package ChMat\ExcelBundle\Tests\Data
 */
class TestFixture {

	public static $header = array(
		'A' => 'Colonne A',
		'B' => 'Colonne B',
		'C' => 'Colonne C',
		'D' => 'Colonne D',
		'E' => 'Colonne E',
	);

	public static $data = array(
		1 => array('A' => 'Colonne A',	'B' => 'Colonne B',	'C' => 'Colonne C',	'D' => 'Colonne D',	'E' => 'Colonne E'),
		2 => array('A' => 'Cellule A2', 'B' => 'Cellule B2', 'C' => '', 'D' => '', 'E' => ''),
		3 => array('A' => '', 'B' => '', 'C' => 'Cellule C3', 'D' => '', 'E' => ''),
		4 => array('A' => '', 'B' => '', 'C' => '', 'D' => 'Cellule D4', 'E' => ''),
		5 => array('A' => '', 'B' => '', 'C' => '', 'D' => '', 'E' => 'Cellule E5'),
		6 => array('A' => '', 'B' => '', 'C' => '', 'D' => '', 'E' => ''),
		7 => array('A' => '', 'B' => '', 'C' => 'Cellule C7', 'D' => '', 'E' => ''),
		8 => array('A' => 'Cellule A8', 'B' => '', 'C' => '', 'D' => '', 'E' => 'Cellule E8'),
		9 => array('A' => '', 'B' => 'Cellule B9', 'C' => '', 'D' => '', 'E' => ''),
		10 => array('A' => '', 'B' => '', 'C' => 'Cellule C10', 'D' => '', 'E' => ''),
	);
	
}