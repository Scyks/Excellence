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

#pragma mark - tear down

	public function tearDown() {
		// try to unlink created xlsx file
		@unlink(dirname(__FILE__) . DIRECTORY_SEPARATOR . 'build' . DIRECTORY_SEPARATOR . 'test.xlsx');

		parent::tearDown();
	}

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

	/**
	 * data provider that returns values that don't match string,
	 * double, float, integer.
	 *
	 * @return array
	 */
	public function dataProviderInvalidValuesForCell() {
		return array(
			array(true),
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
	public function create_DelegateGetSheetForWorkBookReturnsNoInstaceOfSheet_throwsException() {
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
	public function create_createSheetXml_ScheetXmlCreated() {
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

	/**
	 * @test
	 * @group Workbook
	 * @expectedException \LogicException
	 * @expectedExceptionMessage WorkbookDelegate::dataSourceForWorkbookAndSheet have to return an instance of \Excellence\Delegates\DataSource, "NULL" given.
	 */
	public function create_DelegateDataSourceForWorkbookAndSheetReturnsNoInstaceOfDataSource_throwsException() {
		$oMock = $this->getMock('\Test\Excellence\Stub\DataSource', array('dataSourceForWorkbookAndSheet'));
		$oMock
			->expects($this->any())
			->method('dataSourceForWorkbookAndSheet')
			->will($this->returnValue(null))
		;

		$this->makeWorkbook('foo', $oMock)->create();
	}

	/**
	 * @test
	 * @group Workbook
	 * @dataProvider dataProviderInvalidValuesForPositiveIntegers
	 * @expectedException \LogicException
	 * @expectedExceptionMessage DataDelegate::numberOfRowsInSheet have to return an integer value bigger than zero.
	 */
	public function create_DataSourceNumberOfRowsInSheetReturnNotAnInteger_throwsException($value) {
		$oMock = $this->getMock('\Test\Excellence\Stub\DataSource', array('numberOfRowsInSheet'));
		$oMock
			->expects($this->any())
			->method('numberOfRowsInSheet')
			->will($this->returnValue($value))
		;

		$this->makeWorkbook('foo', $oMock)->create();
	}

	/**
	 * @test
	 * @group Workbook
	 * @dataProvider dataProviderInvalidValuesForPositiveIntegers
	 * @expectedException \LogicException
	 * @expectedExceptionMessage DataDelegate::numberOfColumnsInSheet have to return an integer value bigger than zero.
	 */
	public function create_DataSourceNumberOfColumnsInSheetReturnNotAnInteger_throwsException($value) {
		$oMock = $this->getMock('\Test\Excellence\Stub\DataSource', array('numberOfColumnsInSheet'));
		$oMock
			->expects($this->any())
			->method('numberOfColumnsInSheet')
			->will($this->returnValue($value))
		;

		$this->makeWorkbook('foo', $oMock)->create();
	}

	/**
	 * @test
	 * @group Workbook
	 * @dataProvider dataProviderInvalidValuesForCell
	 *
	 */
	public function create_DataSourceValueForRowAndColumnReturnInvalidType_throwsException($value) {
		$oMock = $this->getMock('\Test\Excellence\Stub\DataSource', array('valueForRowAndColumn'));
		$oMock
			->expects($this->any())
			->method('valueForRowAndColumn')
			->will($this->returnValue($value))
		;

		$this->setExpectedException('\LogicException', sprintf('DataDelegate::valueForRowAndColumn have to return a string, float, double or int value, "%s" given.', gettype($value)));

		$this->makeWorkbook('foo', $oMock)->create();
	}

	/**
	 * @test
	 * @group Workbook
	 */
	public function create_createSheetDataXml_ScheetDataXmlCreated() {
		$oWorkbook = $this->makeWorkbook();

		$sCompareXml = '<sheets id="%s" rows="4" columns="3" dimension="A1:C4">'
			. '<row id="1">'
				. '<column id="A1" type="1">row1col1</column>'
				. '<column id="B1" type="2">42</column>'
				. '<column id="C1" type="2">42.34</column>'
			. '</row>'
			. '<row id="2">'
				. '<column id="A2" type="1">row2col1</column>'
				. '<column id="B2" type="2">42</column>'
				. '<column id="C2" type="2">42.34</column>'
			. '</row>'
			. '<row id="3">'
				. '<column id="A3" type="1">row3col1</column>'
				. '<column id="B3" type="2">42</column>'
				. '<column id="C3" type="2">42.34</column>'
			. '</row>'
			. '<row id="4">'
				. '<column id="B4" type="4">SUM(B1:B3)</column>'
				. '<column id="C4" type="4">SUM(C1:C3)</column>'
			. '</row>'
			. '</sheets>'
		;

		$oDom = new \DOMDocument('1.0', 'utf-8');
		$oDom->loadXML(sprintf($sCompareXml, 'sheet1'));

		$oDom2 = new \DOMDocument('1.0', 'utf-8');
		$oDom2->loadXML(sprintf($sCompareXml, 'sheet2'));

		$aCompare = array(
			'sheet1' => $oDom,
			'sheet2' => $oDom2,
		);

		$oWorkbook->create();



		$this->assertAttributeEquals($aCompare, 'aSheetData', $oWorkbook);
	}

	/**
	 * @test
	 * @group Workbook
	 */
	public function create_dataIncludingFurmulas_CalcChainXmlCreated() {
		$oWorkbook = $this->makeWorkbook();

		$sCompareXml = '<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
				. '<c r="B4" i="1"/>'
				. '<c r="C4" i="1"/>'
				. '<c r="B4" i="2"/>'
				. '<c r="C4" i="2"/>'
			. '</calcChain>'
		;

		$oDom = new \DOMDocument('1.0', 'utf-8');
		$oDom->loadXML($sCompareXml);

		$oWorkbook->create();

		$this->assertAttributeEquals($oDom, 'oCalcChain', $oWorkbook);
	}

	/**
	 * @test
	 * @group Workbook
	 */
	public function create_checkChaining_ReturnsSelf() {
		$oWorkbook = $this->makeWorkbook();

		$oCompare = $oWorkbook->create();

		$this->assertSame($oCompare, $oWorkbook);
	}

#pragma mark - save

	/**
	 * @test
	 * @group Workbook
	 */
	public function save_createWorkbook_workbookSaved() {
		$sFilename = dirname(__FILE__) . DIRECTORY_SEPARATOR . 'build' . DIRECTORY_SEPARATOR . 'test.xlsx';
		$oWorkbook = $this->makeWorkbook();
		$oWorkbook
			->create()
			->save($sFilename)
		;

		$this->assertFileExists($sFilename);
	}
}