<?php

/*
 * This file is part of ChMatExcelBundle.
 *
 * (c) Christian Mattart <christian@chmat.be>
 *
 * This source file is subject to the MIT license that is bundled
 * with this source code in the LICENSE file.
 */

namespace ChMat\ExcelBundle\Tests\Service;


use ChMat\ExcelBundle\Service\ExcelReader;
use ChMat\ExcelBundle\Tests\Data\TestFixture;
use PHPUnit_Framework_TestCase;

/**
 * Class ExcelReaderTest
 * @package ChMat\ExcelBundle\Tests\Service
 */
class ExcelReaderTest extends PHPUnit_Framework_TestCase 
{
	/**
	 * @var ExcelReader
	 */
	private $reader;
	
	private function init()
	{
		$this->reader = new ExcelReader();
		$this->reader->load(__DIR__ . '/../data/test_read.xls');
	}
	
	public function testCounts()
	{
		$this->init();
		
		$this->assertEquals(10, $this->reader->getNumberOfRows(), 'Wrong number of rows.');
		$this->assertEquals('E', $this->reader->getNumberOfCols(), 'Wrong number of columns.');
	}

	/**
	 * @depends testCounts
	 */
	public function testData()
	{
		$this->init();
		
		$expectedHeader = TestFixture::$header;
		
		$headers = $this->reader->getColumnHeaders();
		
		$this->assertTrue(is_array($headers), 'getColumnHeaders result should be an array.');
		$this->assertEquals(5, sizeof($headers), 'Headers array row count is not correct.');
		foreach ($expectedHeader as $column => $value)
		{
			$this->assertEquals($value, $headers[$column], sprintf('Column %s header value is incorrect.', $column, $value));
		}
		
		$expectedData = TestFixture::$data;

		$data = $this->reader->read();
		
		$this->assertEquals($expectedData, $data, 'Data array is not formatted as expected.');
	}
}
