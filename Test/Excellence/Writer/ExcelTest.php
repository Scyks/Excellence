<?php
/**
 * @author		  Ronald Marske <r.marske@secu-ring.de>
 * @filesource    Test/Excellence/Writer/AbstractWriterTest.php
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
 *
 */

namespace Test\Excellence\Writer;

use Excellence\Workbook;
use Excellence\Writer\AbstractWriter;
use Test\Excellence\Stub\Writer\Excel;
use Test\TestCase;

/**
 * AbstractWriter Test class
 * @package Test\Excellence\Writer
 * @group AbstractWriter
 */
class ExcelTest extends TestCase {

#pragma mark - creation

	/**
	 * @param Workbook $oWorkbook
	 * @return AbstractWriter
	 */
	private function createWriter(Workbook $oWorkbook = null) {

		if (!$oWorkbook instanceof Workbook)
			$oWorkbook = $this->makeWorkbook();

		return new Excel($oWorkbook);
	}

#pragma mark - dataProvider

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

	/**
	 * data provider that returns column and row number and result cord
	 *
	 * @return array
	 */
	public function dataProviderCellCoordinates() {
		return array(
			array(0, 1, 'A1'),
			array(0, 2, 'A2'),
			array(1, 1, 'B1'),
			array(27, 1, 'AB1'),
			array(27, 11, 'AB11'),
			array(1260, 1, 'AVM1'),
			array(16383 , 1, 'XFD1'),
		);
	}

#pragma mark - construction

	/**
	 * @test
	 */
	public function __construct_givenWorkbook_workbookStored() {
		$oWorkbook = $this->makeWorkbook();

		$oWriter = $this->createWriter($oWorkbook);

		$this->assertAttributeEquals($oWorkbook, 'oWorkbook', $oWriter);
	}

	/**
	 * @test
	 */
	public function getWorkbook_givenWorkbook_workbookReturned() {
		$oWorkbook = $this->makeWorkbook();

		$oWriter = $this->createWriter($oWorkbook);

		$this->assertEquals($oWorkbook, $oWriter->getWorkbook());
	}

#pragma mark - getSheets

	/**
	 * @test
	 */
	public function getNumberOfSheets_notStored_returnsDelegateSheetCount() {
		$oMock = $this->getMock('\Test\Excellence\Stub\DataSource', array('numberOfSheetsInWorkbook'));
		$oMock
			->expects($this->any())
			->method('numberOfSheetsInWorkbook')
			->will($this->returnValue(22))
		;

		$oWorkbook = $this->makeWorkbook('foo', $oMock);
		$this->assertSame(22, $this->createWriter($oWorkbook)->getNumberOfSheets());
	}

	/**
	 * @test
	 */
	public function getNumberOfSheets_noCountStored_storeSheetCount() {

		$oWriter = $this->createWriter();

		$oReflectionClass = new \ReflectionClass($oWriter);
		$oSheet = $oReflectionClass->getProperty('iSheets');
		$oSheet->setAccessible(true);

		$oWriter->getNumberOfSheets();

		$this->assertAttributeEquals(2, 'iSheets', $oWriter);
	}

	/**
	 * @test
	 */
	public function getNumberOfSheets_storedSheetCount_returnstoredSheetCount() {

		$oWriter = $this->createWriter();

		$oReflectionClass = new \ReflectionClass($oWriter);
		$oSheet = $oReflectionClass->getProperty('iSheets');
		$oSheet->setAccessible(true);
		$oSheet->setValue($oWriter, 22);


		$this->assertSame(22, $oWriter->getNumberOfSheets());
	}

#pragma mark - check sheets in Workbook

	/**
	 * @test
	 * @dataProvider dataProviderInvalidValuesForPositiveIntegers
	 * @expectedException \LogicException
	 * @expectedExceptionMessage WorkbookDelegate::numberOfSheetsInWorkbook have to return an integer bigger than zero.
	 */
	public function checkNumberOfSheetsInWorkbook_delegateRetrunsNoIntergerValue_throwException($value) {
		$oMock = $this->getMock('\Test\Excellence\Stub\DataSource', array('numberOfSheetsInWorkbook'));
		$oMock
			->expects($this->any())
			->method('numberOfSheetsInWorkbook')
			->will($this->returnValue($value))
		;

		$oWorkbook = $this->makeWorkbook('foo', $oMock);
		$this->createWriter($oWorkbook)->checkNumberOfSheetsInWorkbook();
	}

	/**
	 * @test
	 */
	public function checkNumberOfSheetsInWorkbook_checkChaining_returnsSelf() {

		$oWriter = $this->createWriter();
		$this->assertSame($oWriter, $oWriter->checkNumberOfSheetsInWorkbook());
	}

#pragma mark - sharedStrings

	/**
	 * @test
	 */
	public function addValueToSharedStrings_stringDoesntExists_returnStorageNumber() {
		$oWriter = $this->createWriter();
		$this->assertSame(0, $oWriter->addValueToSharedStrings('foobar'));
	}

	/**
	 * @test
	 */
	public function addValueToSharedStrings_stringDoesExists_returnStorageNumberOfExisting() {
		$oWriter = $this->createWriter();
		$oWriter->addValueToSharedStrings('foobar');
		$oWriter->addValueToSharedStrings('foobar2');
		$this->assertSame(0, $oWriter->addValueToSharedStrings('foobar'));
	}

#pragma mark - hyperlinks

	/**
	 * @test
	 */
	public function addHyperlink_stringDoesntExists_returnStorageNumber() {
		$oWriter = $this->createWriter();
		$oSheet = $oWriter->getDelegate()->getSheetForWorkBook($oWriter->getWorkbook(), 0);
		$this->assertSame(1, $oWriter->addHyperlink($oSheet, 'foobar'));
	}

	/**
	 * @test
	 */
	public function addHyperlink_stringDoesExists_returnStorageNumberOfExisting() {
		$oWriter = $this->createWriter();
		$oSheet = $oWriter->getDelegate()->getSheetForWorkBook($oWriter->getWorkbook(), 0);
		$oWriter->addHyperlink($oSheet, 'foobar');
		$oWriter->addHyperlink($oSheet, 'foobar2');
		$this->assertSame(1, $oWriter->addHyperlink($oSheet, 'foobar'));
	}

	/**
	 * @test
	 */
	public function hasHyperlinks_noLinks_returnFalse() {
		$oWriter = $this->createWriter();
		$oSheet = $oWriter->getDelegate()->getSheetForWorkBook($oWriter->getWorkbook(), 0);
		$this->assertFalse($oWriter->hasHyperlinks($oSheet));
	}

	/**
	 * @test
	 */
	public function hasHyperlinks_withLinks_returnTrue() {
		$oWriter = $this->createWriter();
		$oSheet = $oWriter->getDelegate()->getSheetForWorkBook($oWriter->getWorkbook(), 0);
		$oWriter->addHyperlink($oSheet, 'foobar');
		$oWriter->addHyperlink($oSheet, 'foobar2');
		$this->assertTrue($oWriter->hasHyperlinks($oSheet));
	}



#pragma mark - createSheetXml

	/**
	 * @test
	 * @expectedException \LogicException
	 * @expectedExceptionMessage WorkbookDelegate::getSheetForWorkBook have to return an instance of \Excellence\Sheet, "NULL" given.
	 */
	public function createSheetXml_DelegateGetSheetForWorkBookReturnsNoInstaceOfSheet_throwsException() {
		$oMock = $this->getMock('\Test\Excellence\Stub\DataSource', array('getSheetForWorkBook'));
		$oMock
			->expects($this->any())
			->method('getSheetForWorkBook')
			->will($this->returnValue(null))
		;

		$oWorkbook = $this->makeWorkbook('foo', $oMock);
		$oWriter = $this->createWriter($oWorkbook);
		$oWriter->createSheetXml();
	}

	/**
	 * @test
	 */
	public function createSheetXml_createSheetXml_ScheetXmlCreated() {
		$oWriter = $this->createWriter();

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


		$this->assertEquals($sCompareXml, $oWriter->createSheetXml());

	}

#pragma mark - createSheetDataXml

	/**
	 * @test
	 * @group Workbook
	 * @expectedException \LogicException
	 * @expectedExceptionMessage WorkbookDelegate::dataSourceForWorkbookAndSheet have to return an instance of \Excellence\Delegates\DataSource, "NULL" given.
	 */
	public function createSheetDataXml_DelegateDataSourceForWorkbookAndSheetReturnsNoInstanceOfDataSource_throwsException() {
		$oMock = $this->getMock('\Test\Excellence\Stub\DataSource', array('dataSourceForWorkbookAndSheet'));
		$oMock
			->expects($this->any())
			->method('dataSourceForWorkbookAndSheet')
			->will($this->returnValue(null))
		;

		$oWorkbook = $this->makeWorkbook('foo', $oMock);
		$oWriter = $this->createWriter($oWorkbook);

		$oWriter->createSheetDataXml($oMock->getSheetForWorkBook($oWorkbook, 0), 0);
	}

	/**
	 * @test
	 * @group Workbook
	 * @dataProvider dataProviderInvalidValuesForPositiveIntegers
	 * @expectedException \LogicException
	 * @expectedExceptionMessage DataDelegate::numberOfRowsInSheet have to return an integer value bigger than zero.
	 */
	public function createSheetDataXml_DataSourceNumberOfRowsInSheetReturnNotAnInteger_throwsException($value) {
		$oMock = $this->getMock('\Test\Excellence\Stub\DataSource', array('numberOfRowsInSheet'));
		$oMock
			->expects($this->any())
			->method('numberOfRowsInSheet')
			->will($this->returnValue($value))
		;

		$oWorkbook = $this->makeWorkbook('foo', $oMock);
		$oWriter = $this->createWriter($oWorkbook);

		$oWriter->createSheetDataXml($oMock->getSheetForWorkBook($oWorkbook, 0), 0);
	}

	/**
	 * @test
	 * @dataProvider dataProviderInvalidValuesForPositiveIntegers
	 * @expectedException \LogicException
	 * @expectedExceptionMessage DataDelegate::numberOfColumnsInSheet have to return an integer value bigger than zero.
	 */
	public function createSheetDataXml_DataSourceNumberOfColumnsInSheetReturnNotAnInteger_throwsException($value) {
		$oMock = $this->getMock('\Test\Excellence\Stub\DataSource', array('numberOfColumnsInSheet'));
		$oMock
			->expects($this->any())
			->method('numberOfColumnsInSheet')
			->will($this->returnValue($value))
		;

		$oWorkbook = $this->makeWorkbook('foo', $oMock);
		$oWriter = $this->createWriter($oWorkbook);

		$oWriter->createSheetDataXml($oMock->getSheetForWorkBook($oWorkbook, 0), 0);
	}

	/**
	 * @test
	 * @dataProvider dataProviderInvalidValuesForCell
	 */
	public function createSheetDataXml_DataSourceValueForRowAndColumnReturnInvalidType_throwsException($value) {
		$oMock = $this->getMock('\Test\Excellence\Stub\DataSource', array('valueForRowAndColumn'));
		$oMock
			->expects($this->any())
			->method('valueForRowAndColumn')
			->will($this->returnValue($value))
		;

		$this->setExpectedException('\LogicException', sprintf('DataDelegate::valueForRowAndColumn have to return a string, float, double or int value, "%s" given.', gettype($value)));

		$oWorkbook = $this->makeWorkbook('foo', $oMock);
		$oWriter = $this->createWriter($oWorkbook);

		$oWriter->createSheetDataXml($oMock->getSheetForWorkBook($oWorkbook, 0), 0);
	}

	/**
	 * @test
	 */
	public function createSheetDataXml_createSheetDataXml_ScheetDataXmlCreated() {
		$oWorkbook = $this->makeWorkbook();

		$sCompareXml = '<?xml version="1.0" encoding="UTF-8"?>'
			. '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">'
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


		$oWriter = $this->createWriter($oWorkbook);
		$oWriter->createSheetDataXml($oWorkbook->getDelegate()->getSheetForWorkBook($oWorkbook, 0), 0);
		$oWriter->createSheetDataXml($oWorkbook->getDelegate()->getSheetForWorkBook($oWorkbook, 1), 1);

		$oReflection = new \ReflectionClass($oWriter);
		$oSheetData = $oReflection->getProperty('aSheetData');
		$oSheetData->setAccessible(true);
		$aSheets = $oSheetData->getValue($oWriter);

		$this->assertXmlStringEqualsXmlString($aSheets['sheet1'], $sCompareXml);
		$this->assertXmlStringEqualsXmlString($aSheets['sheet2'], $sCompareXml);
	}

	/**
	 * @test
	 * @expectedException \InvalidArgumentException
	 * @expectedExceptionMessage MergeableDelegate::mergeByColumnAndRow have to return a merge range as string in format "A1:A2"
	 */
	public function createSheetDataXml_createMergeableReturnWrongMergeRange_throwsException() {

		$oMock = $this->getMock('\Test\Excellence\Stub\MergeableDataSource', array('mergeByColumnAndRow'));
		$oMock
			->expects($this->any())
			->method('mergeByColumnAndRow')
			->will($this->returnValue('foobar'))
		;
		$oWorkbook = $this->makeWorkbook('merge', $oMock);

		$oWriter = $this->createWriter($oWorkbook);
		$oWriter->createSheetDataXml($oWorkbook->getDelegate()->getSheetForWorkBook($oWorkbook, 0), 0);

	}

	/**
	 * @test
	 */
	public function createSheetDataXml_createMergeableSheetDataXml_ScheetDataXmlCreated() {

		$oDataSource = $this->makeMergeDelegate();
		$oWorkbook = $this->makeWorkbook('merge', $oDataSource);

		$sCompareXml = '<?xml version="1.0" encoding="UTF-8"?>'
			. '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">'
			. '<sheetViews>'
			. '<sheetView tabSelected="1" workbookViewId="0"/>'
			. '</sheetViews>'
			. '<sheetData>'
			. '<row r="1">'
			. '<c r="A1" t="s"><v>0</v></c>'
			. '<c r="B1" t="s"><v>1</v></c>'
			. '</row>'
			. '<row r="2">'
			. '<c r="A2" t="s"><v>2</v></c>'
			. '<c r="B2" t="s"><v>3</v></c>'
			. '</row>'
			. '<row r="3">'
			. '<c r="A3" t="s"><v>4</v></c>'
			. '</row>'
			. '<row r="4">'
			. '<c r="A4" t="s"><v>5</v></c>'
			. '</row>'
			. '</sheetData>'
			. '<mergeCells>'
			. '<mergeCell ref="B1:C1"/>'
			. '<mergeCell ref="B2:C3"/>'
			. '<mergeCell ref="A3:A4"/>'
			. '</mergeCells>'
			. '</worksheet>'
		;


		$oWriter = $this->createWriter($oWorkbook);
		$oWriter->createSheetDataXml($oWorkbook->getDelegate()->getSheetForWorkBook($oWorkbook, 0), 0);

		$oReflection = new \ReflectionClass($oWriter);
		$oSheetData = $oReflection->getProperty('aSheetData');
		$oSheetData->setAccessible(true);
		$aSheets = $oSheetData->getValue($oWriter);

		$this->assertXmlStringEqualsXmlString($aSheets['sheet1'], $sCompareXml);
	}

	/**
	 * @test
	 */
	public function createSheetDataXml_createLinkableSheetDataXml_ScheetDataXmlCreated() {

		$oDataSource = $this->makeLinkableDelegate();
		$oWorkbook = $this->makeWorkbook('merge', $oDataSource);

		$sCompareXml = '<?xml version="1.0" encoding="UTF-8"?>'
			. '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">'
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
			. '<hyperlinks>'
			. '<hyperlink ref="A1" r:id="rId1"/>'
			. '<hyperlink ref="A2" r:id="rId1"/>'
			. '<hyperlink ref="A3" r:id="rId1"/>'
			. '</hyperlinks>'
			. '</worksheet>'
		;


		$oWriter = $this->createWriter($oWorkbook);
		$oWriter->createSheetDataXml($oWorkbook->getDelegate()->getSheetForWorkBook($oWorkbook, 0), 0);

		$oReflection = new \ReflectionClass($oWriter);
		$oSheetData = $oReflection->getProperty('aSheetData');
		$oSheetData->setAccessible(true);
		$aSheets = $oSheetData->getValue($oWriter);

		$this->assertXmlStringEqualsXmlString($aSheets['sheet1'], $sCompareXml);
	}

#pragma mark - document relations

	/**
	 * @test
	 */
	public function documentRelationsXml_defaultSettings_XmlGenerated() {

		$sXML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
				. '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
				. '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
				. '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
				. '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
				. '</Relationships>'
		;

		$oWriter = $this->createWriter();

		$this->assertXmlStringEqualsXmlString($sXML, $oWriter->documentRelationsXml());
	}

#pragma mark - document relations

	/**
	 * @test
	 */
	public function documentContentTypesXml_standardContentTypes_XMLGenerated() {

		$sXML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
			. '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
			. '<Default Extension="xml" ContentType="application/xml"/>'
			. '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
			. '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
			. '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'

			. '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
			. '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
			. '</Types>'
		;

		$oWriter = $this->createWriter();

		$this->assertXmlStringEqualsXmlString($sXML, $oWriter->documentContentTypesXml());
	}

	/**
	 * @test
	 */
	public function documentContentTypesXml_relationsWithSharedStrings_XMLGenerated() {

		$sXML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
			. '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
			. '<Default Extension="xml" ContentType="application/xml"/>'
			. '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
			. '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
			. '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
			. '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
			. '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
			. '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
			. '<Override PartName="/xl/worksheet/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
			. '<Override PartName="/xl/worksheet/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
			. '</Types>'
		;

		$oWriter = $this->createWriter();
		$oWriter->createSheetXml();
		$oWriter->createSheetDataXml($oWriter->getDelegate()->getSheetForWorkBook($oWriter->getWorkbook(), 0), 0);
		$oWriter->createSheetDataXml($oWriter->getDelegate()->getSheetForWorkBook($oWriter->getWorkbook(), 1), 1);

		$this->assertXmlStringEqualsXmlString($sXML, $oWriter->documentContentTypesXml());
	}

	/**
	 * @test
	 */
	public function workbookAppXml_relationsWithSheets_XMLGenerated() {

		$sXML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
			. '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
			. '<Application>Excellence</Application>'
			. '<DocSecurity>0</DocSecurity>'
			. '<ScaleCrop>false</ScaleCrop>'
			. '<HeadingPairs>'
			. '<vt:vector size="2" baseType="variant">'
			. '<vt:variant>'
			. '<vt:lpstr>Arbeitsblätter</vt:lpstr>'
			. '</vt:variant>'
			. '<vt:variant>'
			. '<vt:i4>2</vt:i4>'
			. '</vt:variant>'
			. '</vt:vector>'
			. '</HeadingPairs>'
			. '<TitlesOfParts>'
			.'<vt:vector size="2" baseType="lpstr">'
			. '<vt:lpstr>Sheet 1</vt:lpstr>'
			. '<vt:lpstr/>'
			. '</vt:vector>'
			. '</TitlesOfParts>'
			. '<Company>Excellence</Company>'
			. '<LinksUpToDate>false</LinksUpToDate>'
			. '<SharedDoc>false</SharedDoc>'
			. '<HyperlinksChanged>false</HyperlinksChanged>'
			. '<AppVersion>14.0300</AppVersion>'
			. '</Properties>'
		;

		$oWriter = $this->createWriter();
		$oWriter->createSheetXml();
		$oWriter->createSheetDataXml($oWriter->getDelegate()->getSheetForWorkBook($oWriter->getWorkbook(), 0), 0);
		$oWriter->createSheetDataXml($oWriter->getDelegate()->getSheetForWorkBook($oWriter->getWorkbook(), 1), 1);

		$this->assertXmlStringEqualsXmlString($sXML, $oWriter->workbookAppXml());
	}

	/**
	 * @test
	 */
	public function workbookCoreXml_defaultData_XMLGenerated() {

		$sXML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
			. '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
			// @ todo insert creators name - metata delegate
			. '<dc:creator>creator name</dc:creator>'
			. '<cp:lastModifiedBy>Name of last modified</cp:lastModifiedBy>'
			. '<dcterms:created xsi:type="dcterms:W3CDTF">' . date('Y-m-d\TH:i:s\Z') . '</dcterms:created>'
			. '<dcterms:modified xsi:type="dcterms:W3CDTF">' . date('Y-m-d\TH:i:s\Z') . '</dcterms:modified>'
			. '</cp:coreProperties>'
		;

		$oWriter = $this->createWriter();

		$this->assertXmlStringEqualsXmlString($sXML, $oWriter->workbookCoreXml());
	}

	/**
	 * @test
	 */
	public function workbookRelationsXml_relationsWithSheets_XMLGenerated() {

		$sXML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
			. '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
			. '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
			. '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>'
			. '<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
			. '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
			. '</Relationships>'
		;

		$oWriter = $this->createWriter();
		$oWriter->createSheetXml();
		$oWriter->createSheetDataXml($oWriter->getDelegate()->getSheetForWorkBook($oWriter->getWorkbook(), 0), 0);
		$oWriter->createSheetDataXml($oWriter->getDelegate()->getSheetForWorkBook($oWriter->getWorkbook(), 1), 1);

		$this->assertXmlStringEqualsXmlString($sXML, $oWriter->workbookRelationsXml());
	}

	/**
	 * @test
	 */
	public function workbookStylesXML_standardStyles_XMLGenerated() {

		$sXML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
			. '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">'
			. '<fonts count="1" x14ac:knownFonts="1">'
			. '<font>'
			. '<sz val="12"/>'
			. '<name val="Calibri"/>'
			. '<color rgb="FF000000"/>'
			. '</font>'
			. '</fonts>'
			. '<fills count="2">'
			. '<fill>'
			. '<patternFill patternType="none"/>'
			. '</fill>'
			. '<fill>'
			. '<patternFill patternType="gray125"/>'
			. '</fill>'
			. '</fills>'
			. '<borders count="1">'
			. '<border/>'
			. '</borders>'
			. '<cellStyleXfs count="1">'
			. '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>'
			. '</cellStyleXfs>'
			. '<cellXfs count="1">'
			. '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" shrinkToFit="true" wrapText="true"/>'
			. '</cellXfs>'
			. '<cellStyles count="1">'
			. '<cellStyle name="Standard" xfId="0" builtinId="0"/>'
			. '</cellStyles>'
			. '<dxfs count="0"/>'
			. '<tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleMedium4"/>'
			. '</styleSheet>'
		;

		$oWriter = $this->createWriter();
		$oWriter->createSheetXml();
		$oWriter->createSheetDataXml($oWriter->getDelegate()->getSheetForWorkBook($oWriter->getWorkbook(), 0), 0);
		$oWriter->createSheetDataXml($oWriter->getDelegate()->getSheetForWorkBook($oWriter->getWorkbook(), 1), 1);

		$this->assertXmlStringEqualsXmlString($sXML, $oWriter->workbookStylesXML());
	}

	/**
	 * @test
	 */
	public function workbookStylesXML_customizedStyles_XMLGenerated() {

		$sXML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
				. '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">'
				. '<fonts count="3" x14ac:knownFonts="3">'
				. '<font>'
				. '<sz val="12"/>'
				. '<name val="Calibri"/>'
				. '<color rgb="FF000000"/>'
				. '</font>'
				. '<font>'
				. '<sz val="20"/>'
				. '<name val="Arial"/>'
				. '<color rgb="FF000000"/>'
				. '<b/>'
				. '</font>'
				. '<font>'
				. '<b/>'
				. '<i/>'
				. '<u/>'
				. '</font>'
				. '</fonts>'
				. '<fills count="4">'
				. '<fill>'
				. '<patternFill patternType="none"/>'
				. '</fill>'
				. '<fill>'
				. '<patternFill patternType="gray125"/>'
				. '</fill>'
				. '<fill>'
				. '<patternFill patternType="solid">'
				. '<fgColor rgb="FFFCFAE4"/>'
				. '</patternFill>'
				. '</fill>'
				. '<fill>'
				. '<patternFill patternType="solid">'
				. '<fgColor rgb="FFFBF6BE"/>'
				. '</patternFill>'
				. '</fill>'
				. '</fills>'
				. '<borders count="3">'
				. '<border/>'
				. '<border>'
				. '<left style="thin">'
				. '<color rgb="FF0949E9"/>'
				. '</left>'
				. '<right style="thin">'
				. '<color rgb="FF0949E9"/>'
				. '</right>'
				. '<top style="thin">'
				. '<color rgb="FF0949E9"/>'
				. '</top>'
				. '<bottom style="thin">'
				. '<color rgb="FF0949E9"/>'
				. '</bottom>'
				. '</border>'
				. '<border>'
				. '<left style="thin">'
				. '<color rgb="FF333333"/>'
				. '</left>'
				. '<right style="thin">'
				. '<color rgb="FF333333"/>'
				. '</right>'
				. '<top style="thin">'
				. '<color rgb="FF333333"/>'
				. '</top>'
				. '<bottom style="thin">'
				. '<color rgb="FF333333"/>'
				. '</bottom>'
				. '</border>'
				. '</borders>'
				. '<cellStyleXfs count="1">'
				. '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>'
				. '</cellStyleXfs>'
				. '<cellXfs count="4">'
				. '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" shrinkToFit="true" wrapText="true"/>'
				. '<xf xfId="0" numFmtId="0" fontId="1" fillId="0" borderId="1" shrinkToFit="true" wrapText="true">'
				. '<alignment horizontal="center" vertical="center"/>'
				. '</xf>'
				. '<xf xfId="0" numFmtId="0" fontId="0" fillId="2" borderId="1" shrinkToFit="true" wrapText="true">'
				. '</xf>'
				. '<xf xfId="0" numFmtId="0" fontId="2" fillId="3" borderId="2" shrinkToFit="true" wrapText="true">'
				. '<alignment horizontal="center" vertical="center"/>'
				. '</xf>'
				. '</cellXfs>'
				. '<cellStyles count="1">'
				. '<cellStyle name="Standard" xfId="0" builtinId="0"/>'
				. '</cellStyles>'
				. '<dxfs count="0"/>'
				. '<tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleMedium4"/>'
				. '</styleSheet>'
		;

		
		$oStleableDelegate = $this->makeStylableDelegate();
		$oWorkbook = $this->makeWorkbook('workbook', $oStleableDelegate);
		$oWriter = $this->createWriter($oWorkbook);
		$oWriter->createSheetXml();
		$oWriter->createSheetDataXml($oWriter->getDelegate()->getSheetForWorkBook($oWriter->getWorkbook(), 0), 0);
		$oWriter->createSheetDataXml($oWriter->getDelegate()->getSheetForWorkBook($oWriter->getWorkbook(), 1), 1);

		$this->assertXmlStringEqualsXmlString($sXML, $oWriter->workbookStylesXML());
	}

	/**
	 * @test
	 */
	public function sharedStringsXml_standardStyles_XMLGenerated() {

		$sXML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
				. '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">'
				. '<si>'
				. '<t>row1col1</t>'
				. '</si>'
				. '<si>'
				. '<t>row2col1</t>'
				. '</si>'
				. '<si><t>row3col1</t>'
				. '</si>'
				. '</sst>'
		;

		$oWriter = $this->createWriter();
		$oWriter->createSheetXml();
		$oWriter->createSheetDataXml($oWriter->getDelegate()->getSheetForWorkBook($oWriter->getWorkbook(), 0), 0);
		$oWriter->createSheetDataXml($oWriter->getDelegate()->getSheetForWorkBook($oWriter->getWorkbook(), 1), 1);

		$this->assertXmlStringEqualsXmlString($sXML, $oWriter->sharedStringsXml());
	}

	/**
	 * @test
	 */
	public function worksheetRelationsXml_includingHyperlinks_XMLGenerated() {

		$sXML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
				. '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
				. '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="http://google.de" TargetMode="External"/>'
				. '</Relationships>'
		;

		$oLinkableDelegate = $this->makeLinkableDelegate();
		$oWorkbook = $this->makeWorkbook('workbook', $oLinkableDelegate);
		$oWriter = $this->createWriter($oWorkbook);

		$oWriter->createSheetXml();
		$oSheet = $oWriter->getDelegate()->getSheetForWorkBook($oWriter->getWorkbook(), 0);
		$oWriter->createSheetDataXml($oSheet, 0);


		$this->assertXmlStringEqualsXmlString($sXML, $oWriter->worksheetRelationsXml($oSheet));
	}

#pragma mark - create file

	/**
	 * @test
	 * @expectedException \LogicException
	 * @expectedExceptionMessage Excellence could not create Excel workbook.
	 */
	public function  saveToFile_fileExists_ThrowsException() {
		$sFilename = dirname(dirname(__FILE__)) . DIRECTORY_SEPARATOR . 'build' . DIRECTORY_SEPARATOR . '.gitignore';
		$oLinkableDelegate = $this->makeLinkableDelegate();
		$oWorkbook = $this->makeWorkbook('workbook', $oLinkableDelegate);
		$oWriter = $this->createWriter($oWorkbook);

		$oWriter->saveToFile($sFilename);

	}

	/**
	 * @test
	 */
	public function  saveToFile_createExcelFile_ExcelFileCreated() {
		$sFilename = dirname(dirname(__FILE__)) . DIRECTORY_SEPARATOR . 'build' . DIRECTORY_SEPARATOR . 'test.xlsx';

		$oLinkableDelegate = $this->makeLinkableDelegate();
		$oWorkbook = $this->makeWorkbook('workbook', $oLinkableDelegate);
		$oWriter = $this->createWriter($oWorkbook);

		$oWriter->saveToFile($sFilename);

		$this->assertFileExists($sFilename);

		@unlink($sFilename);
	}
}