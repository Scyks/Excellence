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
			array(array()),
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

#pragma mark - delegate

	/**
	 * @test
	 * @group Workbook
	 */
	public function getDelegate_delegateDefined_ReturnsDelegate() {
		$oDelegate = $this->makeDelegate();
		$oWorkBook = $this->makeWorkbook('id', $oDelegate);

		$this->assertSame($oDelegate, $oWorkBook->getDelegate());
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

		$sCompareXml = '<?xml version="1.0" encoding="UTF-8"?>' . "\n"
			. '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
				. '<fileVersion appName="xl" lastEdited="5" lowestEdited="5" rupBuild="23206"/>'
				. '<workbookPr showInkAnnotation="0" autoCompressPictures="0"/>'
				. '<bookViews>'
					. '<workbookView xWindow="0" yWindow="0" windowWidth="25600" windowHeight="14460" tabRatio="500"/>'
				. '</bookViews>'
				. '<sheets>'
					. '<sheet name="Sheet 1" sheetId="1" r:id="rId1"/>'
					. '<sheet name="Sheet 2" sheetId="2" r:id="rId2"/>'
				. '</sheets>'
				. '<calcPr calcId="140000" concurrentCalc="0"/>'
				. '<extLst>'
					. '<ext xmlns:mx="http://schemas.microsoft.com/office/mac/excel/2008/main" uri="{7523E5D3-25F3-A5E0-1632-64F254C22452}">'
						. '<mx:ArchID Flags="2"/>'
					. '</ext>'
				. '</extLst>'
			. "</workbook>";


		$oWorkbook->create();

		$this->assertAttributeEquals($sCompareXml, 'sWorkbook', $oWorkbook);

	}

	/**
	 * @test
	 * @group Workbook
	 * @expectedException \LogicException
	 * @expectedExceptionMessage WorkbookDelegate::dataSourceForWorkbookAndSheet have to return an instance of \Excellence\Delegates\DataSource, "NULL" given.
	 */
	public function create_DelegateDataSourceForWorkbookAndSheetReturnsNoInstanceOfDataSource_throwsException() {
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

		$sCompareXml = '<?xml version="1.0" encoding="UTF-8"?>'
		. '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">'
			. '<dimension ref="A1:D4"/>'
			. '<sheetViews>'
				. '<sheetView tabSelected="1" workbookViewId="0"/>'
			. '</sheetViews>'
			. '<sheetData>'
				. '<row r="1">'
					. '<c r="A1" t="s"><v>0</v></c>'
					. '<c r="B1" t="n"><v>42</v></c>'
					. '<c r="C1" t="n"><v>42.34</v></c>'
					. '<c r="D1" t="b"><v>1</v></c>'
				. '</row>'
				. '<row r="2">'
					. '<c r="A2" t="s"><v>1</v></c>'
					. '<c r="B2" t="n"><v>42</v></c>'
					. '<c r="C2" t="n"><v>42.34</v></c>'
					. '<c r="D2" t="b"><v>0</v></c>'
				. '</row>'
				. '<row r="3">'
					. '<c r="A3" t="s"><v>2</v></c>'
					. '<c r="B3" t="n"><v>42</v></c>'
					. '<c r="C3" t="n"><v>42.34</v></c>'
					. '<c r="D3" t="b"><v>1</v></c>'
				. '</row>'
				. '<row r="4">'
					. '<c r="B4"><f>SUM(B1:B3)</f></c>'
					. '<c r="C4"><f>SUM(C1:C3)</f></c>'
				. '</row>'
			. '</sheetData>'
		. '</worksheet>'
		;


		$oWorkbook->create();

		$oReflection = new \ReflectionClass($oWorkbook);
		$oSheetData = $oReflection->getProperty('aSheetData');
		$oSheetData->setAccessible(true);
		$aSheets = $oSheetData->getValue($oWorkbook);

		$this->assertXmlStringEqualsXmlString($aSheets['sheet1'], $sCompareXml);
		$this->assertXmlStringEqualsXmlString($aSheets['sheet2'], $sCompareXml);
	}

	/**
	 * @test
	 * @group Workbook
	 */
	public function create_SharedStrings_SharedStringsCreated() {
		$oWorkbook = $this->makeWorkbook();

		$aCompare = array(
			'row1col1' => 0,
			'row2col1' => 1,
			'row3col1' => 2,
		);

		$oWorkbook->create();

		$this->assertAttributeEquals($aCompare, 'aSharedStrings', $oWorkbook);
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
		$iTime = microtime(true);
		$sFilename = dirname(__FILE__) . DIRECTORY_SEPARATOR . 'build' . DIRECTORY_SEPARATOR . 'test.xlsx';
		$oWorkbook = $this->makeWorkbook();
		$oWorkbook
			->create()
			->save($sFilename)
		;

		$this->assertFileExists($sFilename);

		//echo "\n" . (microtime(true)-$iTime) . "\n";
		@unlink(dirname(__FILE__) . DIRECTORY_SEPARATOR . 'build' . DIRECTORY_SEPARATOR . 'test.xlsx');
	}


	/**
	 * @test
	 * @group Workbook
	 */
	public function save_createWorkbookPerformanceTest_SaveWorkBookIncludingNColumnsAnd4RowsUnter20Seconds() {

		//$this->markTestSkipped('only for performance optimization');

		$sFilename = dirname(__FILE__) . DIRECTORY_SEPARATOR . 'build' . DIRECTORY_SEPARATOR . 'test.xlsx';

		$oDataSource = new \Test\Excellence\Stub\PerformanceDataSource();
		$iTime = microtime(true);

		$oWorkbook = $this->makeWorkbook('performance', $oDataSource);
		$oWorkbook
			->create()
			->save($sFilename)
		;

		$this->assertLessThan(0.36, (microtime(true)-$iTime));

	}
}

