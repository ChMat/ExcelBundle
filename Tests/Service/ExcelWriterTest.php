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


use ChMat\ExcelBundle\Service\ExcelWriter;
use ChMat\ExcelBundle\Tests\Data\TestFixture;

class ExcelWriterTest extends \PHPUnit_Framework_TestCase {

	/**
	 * This only tests that the ExcelWriter->saveFile() function writes a file to disk.
	 * 
	 * @throws \ChMat\ExcelBundle\Exception\FileExistsException
	 * @throws \Exception
	 */
	public function testWriteXls()
	{
		$file = new ExcelWriter();
		$file->createDocument('Test Worksheet', 'This is a Test Worksheet', 'PHPUnit Instance', 'en_US');
		$file->fillSheet(TestFixture::$data, false);
		
		$saveTo = __DIR__ . '/../Data/test_write.xls';
		
		if (file_exists($saveTo))
		{
			@unlink($saveTo);
			$this->assertFileNotExists($saveTo, sprintf('Test file %s could not be removed before write test.', $saveTo));
		}
		
		$file->saveFile($saveTo, 'Excel5');
		
		$this->assertFileExists($saveTo, sprintf('Apparently, file %s could not be written during write test.', $saveTo));
		
	}
}
