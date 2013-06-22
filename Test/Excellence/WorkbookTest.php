<?php
/**
 * @author        Ronald Marske <scyks@ceow.de>
 * @filesource    ExcellenceTest.php
 *
 * @copyright     Copyright (c) 2013 Ronald Marske, All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions
 * are met:
 *
 *     * Redistributions of source code must retain the above copyright
 *       notice, this list of conditions and the following disclaimer.
 *
 *     * Redistributions in binary form must reproduce the above copyright
 *       notice, this list of conditions and the following disclaimer in
 *       the documentation and/or other materials provided with the
 *       distribution.
 *
 *     * Neither the name of Ronald Marske nor the names of his
 *       contributors may be used to endorse or promote products derived
 *       from this software without specific prior written permission.
 *
 * THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
 * "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
 * LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS
 * FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE
 * COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT,
 * INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING,
 * BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
 * LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
 * CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT
 * LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN
 * ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
 * POSSIBILITY OF SUCH DAMAGE.
 */

namespace Test\Excellence;
use Excellence\Workbook;
use Test\Excellence\Stub\DataSource;
use Excellence\Delegates\WorkbookDelegate;

class WorkbookTest extends \PHPUnit_Framework_TestCase {

#pragma mark - creations

	/**
	 * @param string $sIdentifier
	 * @param WorkbookDelegate $oDelegate
	 *
	 * @return Workbook
	 */
	public function makeWorkbook($sIdentifier = 'workbook', WorkbookDelegate $oDelegate = null) {

		if (null == $oDelegate) {
			$oDelegate  = $this->makeDelegate();
		}

		return new Workbook($sIdentifier, $oDelegate);
	}

	/**
	 * @return DataSource
	 */
	public function makeDelegate() {
		return new DataSource();
	}


#pragma mark - dataProvider

	/**
	 * data provider that returns values that don't match or could
	 * type casted to positive integer values.
	 *
	 * @return array
	 */
	public function dataProviderInvalidValuesForPositiveIntegers() {
		return array(
			array(0),
			array(-2),
			array('0'),
			array('0.9'),
			array('-1'),
			array('test'),
			array(array()),
			array(false),
		);
	}

#pragma mark - construction

	/**
	 * @test
	 * @group Workbook
	 * @expectedException \InvalidArgumentException
	 * @expectedExceptionMessage Workbook identifier have to be a non empty string value.
	 */
	public function __construct_provideEmptyIdentifier_throwsException() {
		$this->makeWorkbook('');
	}


	/**
	 * @test
	 * @group Workbook
	 */
	public function __construct_provideStringIdentifier_identifierStored() {
		$oWorkBook = $this->makeWorkbook('foo');
		$this->assertAttributeSame('foo', 'sIdentifier', $oWorkBook);
	}

	/**
	 * @test
	 * @group Workbook
	 */
	public function __construct_checkStoringDelegate_delegateStored() {
		$oDataSource = $this->makeDelegate();
		$oWorkBook = $this->makeWorkbook('foo', $oDataSource);
		$this->assertAttributeSame($oDataSource, 'oDelegate', $oWorkBook);
	}

#pragma mark - identifier

	/**
	 * @test
	 * @group Workbook
	 */
	public function getIdentifier_modifiedIdentifierVariable_returnsStroredVariable() {
		$oWorkBook = $this->makeWorkbook();

		$oReflectionClass = new \ReflectionClass($oWorkBook);

		$oProperty = $oReflectionClass->getProperty('sIdentifier');
		$oProperty->setAccessible(true);
		$oProperty->setValue($oWorkBook, 'foobar');

		$this->assertSame('foobar', $oWorkBook->getIdentifier());
	}

	/**
	 * @test
	 * @group Workbook
	 */
	public function getIdentifier_constructionIdentifier_returnsStroredVariable() {
		$oWorkBook = $this->makeWorkbook();
		$this->assertSame('workbook', $oWorkBook->getIdentifier());
	}

#pragma mark - create

	/**
	 * @test
	 * @group Workbook
	 * @dataProvider dataProviderInvalidValuesForPositiveIntegers
	 * @expectedException \LogicException
	 * @expectedExceptionMessage WorkbookDelegate::numberOfSheetsInWorkbook have to return an integer bigger than zero.
	 */
	public function create_DelegateReturnsWrongSheetNumber_throwsException($value) {
		$oMock = $this->getMock('\Test\Excellence\Stub\DataSource', array('numberOfSheetsInWorkbook'));
		$oMock
			->expects($this->any())
			->method('numberOfSheetsInWorkbook')
			->will($this->returnValue($value))
		;

		$this->makeWorkbook('foo', $oMock)->create();
	}

	/**
	 * @test
	 * @group Workbook
	 * @expectedException \LogicException
	 * @expectedExceptionMessage WorkbookDelegate::getSheetForWorkBook have to return an instance of \Excellence\Sheet, "NULL" given.
	 */
	public function create_DelegateReturnsNoInstaceOfSheet_throwsException() {
		$oMock = $this->getMock('\Test\Excellence\Stub\DataSource', array('getSheetForWorkBook'));
		$oMock
			->expects($this->any())
			->method('getSheetForWorkBook')
			->will($this->returnValue(null))
		;

		$this->makeWorkbook('foo', $oMock)->create();
	}

	/**
	 * @test
	 * @group Workbook
	 */
	public function create_createSheetData_ScheetXmlDataCreated() {
		$oWorkbook = $this->makeWorkbook();

		$sCompareXml = '<sheets count="2">'
			. '<sheet id="sheet1">Sheet 1</sheet>'
			. '<sheet id="sheet2">Sheet 2</sheet>'
			. '</sheets>'
		;

		$oDom = new \DOMDocument('1.0', 'utf-8');
		$oDom->loadXML($sCompareXml);

		$oWorkbook->create();

		$this->assertAttributeEquals($oDom, 'oSheets', $oWorkbook);
	}
}