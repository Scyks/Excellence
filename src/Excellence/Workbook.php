<?php
/**
 * @author        Ronald Marske <scyks@ceow.de>
 * @filesource    src/Excellence/Workbook.php
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

namespace Excellence;

use Excellence\Delegates\DataDelegate;
use Excellence\Delegates\WorkbookDelegate;
use Excellence\Sheet;

/**
 * Class Workbook
 *
 * @package Excellence
 */
class Workbook {


#pragma mark - member variables

	/**
	 * workbook identifier
	 * @var string
	 */
	private $sIdentifier;

	/**
	 * Data source instance
	 * @var WorkbookDelegate
	 */
	private $oDelegate;

	/**
	 * contains dom document including sheets
	 *
	 * @var array
	 */
	private $aSheets = array();

	/**
	 * contains number of sheets in document
	 * @var int
	 */
	private $iSheets = 0;

	/**
	 * Contains workbook XML file as string because it's faster than Domdocument
	 * @var string
	 */
	private $sWorkbook = '';

	/**
	 * contains an array [sheet identifier] => DomDocument including sheet data.
	 *
	 * @var array
	 */
	private $aSheetData;

	/**
	 * contains calc chain xml nodes as strings
	 *
	 * @var string
	 */
	private $sCalcChain = '';

	/**
	 * contains array with shared strings
	 * string => id
	 * @var array
	 */
	private $aSharedStrings = array();

#pragma mark - construction

	/**
	 * create a new workbook
	 *
	 * @param string $sIdentifier
	 * @param WorkbookDelegate $oDelegate
	 * @throws \InvalidArgumentException
	 */
	public function __construct($sIdentifier, WorkbookDelegate $oDelegate) {

		if (empty($sIdentifier)) {
			throw new \InvalidArgumentException('Workbook identifier have to be a non empty string value.');
		}

		$this->sIdentifier = (string) $sIdentifier;
		$this->oDelegate = $oDelegate;
	}

#pragma mark - delegation

	/**
	 * return defined delegate object
	 *
	 * @return WorkbookDelegate
	 */
	public function getDelegate() {
		return $this->oDelegate;
	}

#pragma mark - identifier

	/**
	 * return workbook identifier to identify a workbook. This
	 * should be unique in an application.
	 *
	 * @return string
	 */
	public function getIdentifier() {
		return $this->sIdentifier;
	}

#pragma mark - data handling

	/**
	 * this method will create all needed XML Documents to render with
	 * XSLT stylesheet files.
	 *
	 * @throws \LogicException
	 */
	public function create() {

		// get number of sheets from delegate
		$this->iSheets = (int) $this->getDelegate()->numberOfSheetsInWorkbook($this);

		// make sure that there is minimum one sheet
		if (0 >= $this->iSheets) {
			throw new \LogicException('WorkbookDelegate::numberOfSheetsInWorkbook have to return an integer bigger than zero.');
		}

		// create Sheet Xml Data
		$this->createSheetXml();

		return $this;

	}

	/**
	 * creates DomDocument for Excel sheets that would be used to create an XML
	 * by using XSLT Stylesheet. Of course it is possible to create real OpenXML
	 * Document, but with an XSLT it is easier to handle namespaces, format
	 * changes and further development s of this library.
	 *
	 * @throws \LogicException
	 */
	private function createSheetXml() {

		$this->sWorkbook = '<?xml version="1.0" encoding="UTF-8"?>' . "\n"
			. '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><fileVersion appName="xl" lastEdited="5" lowestEdited="5" rupBuild="23206"/><workbookPr showInkAnnotation="0" autoCompressPictures="0"/>'
			. '<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="25600" windowHeight="14460" tabRatio="500"/></bookViews>'
			. '<sheets>';

			// iterate sheets
		for($iSheet = 0; $iSheet < $this->iSheets; $iSheet++) {

			/** @var Sheet $oSheet */
			$oSheet = $this->getDelegate()->getSheetForWorkBook($this, $iSheet);

			// make sure that oSheet is instanceof sheet
			if (!$oSheet instanceof Sheet) {
				throw new \LogicException(sprintf('WorkbookDelegate::getSheetForWorkBook have to return an instance of \Excellence\Sheet, "%s" given.', gettype($oSheet)));
			}

			$this->aSheets[] = $oSheet;

			$this->sWorkbook .= '<sheet name="' . (($oSheet->hasName()) ? $oSheet->getName() : 'Sheet ' . ($iSheet + 1)) . '" sheetId="' . ($iSheet + 1) . '" r:id="rId' . ($iSheet + 1) . '"/>';

			// create sheet data by sheet
			$this->createSheetDataXml($oSheet, $iSheet + 1);
		}

		$this->sWorkbook .= '</sheets><calcPr calcId="140000" concurrentCalc="0"/><extLst><ext xmlns:mx="http://schemas.microsoft.com/office/mac/excel/2008/main" uri="{7523E5D3-25F3-A5E0-1632-64F254C22452}"><mx:ArchID Flags="2"/></ext></extLst></workbook>';


	}

	/**
	 * generate values for given sheet.
	 *
	 * @param Sheet $oSheet
	 * @param int $iSheet
	 *
	 * @throws \LogicException
	 */
	private function createSheetDataXml(Sheet $oSheet, $iSheet) {

		/** @var DataDelegate $oDataSource */
		$oDataSource = $this->getDelegate()->dataSourceForWorkbookAndSheet($this, $oSheet);

		if (!$oDataSource instanceof DataDelegate) {
			throw new \LogicException('WorkbookDelegate::dataSourceForWorkbookAndSheet have to return an instance of \Excellence\Delegates\DataSource, "NULL" given.');
		}

		// get rows
		$iRows = (int) $oDataSource->numberOfRowsInSheet($this, $oSheet);

		// throw exception if retrieving other value than positive inter bigger than zero
		if (0 >= $iRows) {
			throw new \LogicException('DataDelegate::numberOfRowsInSheet have to return an integer value bigger than zero.');
		}

		// get number of columns
		$iColumns = (int) $oDataSource->numberOfColumnsInSheet($this, $oSheet);

		if (0 >= $iColumns) {
			throw new \LogicException('DataDelegate::numberOfColumnsInSheet have to return an integer value bigger than zero.');
		}


		// Dimension Vars
		$sDimensionFrom = null;
		$sDimensionTo = null;

		// create workbook xml
		$sWorkbook = '<?xml version="1.0" encoding="UTF-8"?>' . "\n"
		. '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">'
			. '<dimension ref="%s"/>'
			. '<sheetViews>'
				. '<sheetView tabSelected="1" workbookViewId="0"/>'
			. '</sheetViews>'
			.'<sheetData>'
		;


		// sheet loop
		for($iRow = 0; $iRow < $iRows; $iRow++) {

			// create row
			$sWorkbook .= '<row r="' . ($iRow + 1) . '">';

			for($iColumn = 0; $iColumn < $iColumns; $iColumn++) {

				// get value for cell
				$value = $oDataSource->valueForRowAndColumn($this, $oSheet, $iRow, $iColumn);

				// Excel coordinate
				$sCord = $this->getCoordinatesByColumnAndRow($iColumn, $iRow + 1);

				// set start dimension
				if (null == $sDimensionFrom) {
					$sDimensionFrom = $sCord;
				}

				// set end dimension
				if (null == $sDimensionTo || $sDimensionTo < $sCord) {
					$sDimensionTo = $sCord;
				}

				// return if value is empty
				if (null === $value) continue;

				// value type
				$iType = gettype($value);

				// make sure we get an allowed data type
				if (!in_array($iType, array('string', 'integer', 'float', 'double', 'boolean'))) {
					throw new \LogicException(sprintf('DataDelegate::valueForRowAndColumn have to return a string, float, double or int value, "%s" given.', $iType));
				}


				// function or formula
				if ('string' == $iType && '=' == substr($value, 0, 1)) {

					// add value to calchain
					$this->addColumnToCalcChain($sCord, $iSheet);

					// add value to column
					$sWorkbook .= '<c r="' . $sCord . '"><f>' . substr($value, 1) . '</f></c>';

				// string
				} elseif('string' == $iType) {
					$iNum = $this->addValueToSharedStrings($value);

					// add value to column
					$sWorkbook .= '<c r="' . $sCord . '" t="s"><v>' . $iNum . '</v></c>';

				// boolean
				} elseif ('boolean' == $iType) {

					// add value to column
					$sWorkbook .= '<c r="' . $sCord . '" t="b"><v>' . (int) $value . '</v></c>';

				// number
				} else {

					// add value to column
					$sWorkbook .= '<c r="' . $sCord . '" t="n"><v>' . $value . '</v></c>';
				}

				// add column to row
			}

			$sWorkbook .= '</row>';
		}

		$sWorkbook .= '</sheetData></worksheet>';

		$this->aSheetData[$oSheet->getIdentifier()] = sprintf($sWorkbook, $sDimensionFrom . ':' . $sDimensionTo);
	}

	/**
	 * calculate by given column and row index the Excel coordinate.
	 * This method is inspired by an Answer of stackoverflow.
	 *
	 * @see http://stackoverflow.com/questions/3302857/algorithm-to-get-the-excel-like-column-name-of-a-number
	 * @param integer$iColumn
	 * @param integer $iRow
	 *
	 * @return string
	 */
	private function getCoordinatesByColumnAndRow($iColumn, $iRow) {

		for($sReturn = ""; $iColumn >= 0; $iColumn = intval($iColumn / 26) - 1) {
			$sReturn = chr($iColumn%26 + 0x41) . $sReturn;
		}

		return $sReturn . $iRow;

	}

	/**
	 * add a column to calc chain. if there is no calc chain document created
	 * this method will create one. To work with strings will increase memory
	 * and it is faster than working with DomDocument or arrays.
	 *
	 * @param string $sCord
	 * @param int $iSheet
	 */
	private function addColumnToCalcChain($sCord, $iSheet) {

		$this->sCalcChain .= '<c r="' . $sCord . '" i="' . $iSheet . '"/>';

	}

	/**
	 * add a string to shared strings and returns index of this value
	 * @param string $value
	 *
	 * @return int
	 */
	private function addValueToSharedStrings($value) {

		// check if value currently exists and return index
		if (array_key_exists($value, $this->aSharedStrings)) {
			return $this->aSharedStrings[$value];
		}

		// get current count
		$iNum = count($this->aSharedStrings);

		// add value to shared strings table
		$this->aSharedStrings[$value] = $iNum;

		// return current index;
		return $iNum;


	}

#pragma mark - saving and creating officeOpenDocument files

	/**
	 * this method will save an xlsx file
	 * @param string $sFilename
	 *
	 * @throws \LogicException
	 */
	public function save($sFilename) {

		// create zip file
		$oZip  = new \ZipArchive;
		$oRes = $oZip->open($sFilename, \ZipArchive::CREATE | \ZipArchive::OVERWRITE);

		if (true !== $oRes) {
			throw new \LogicException('Excellence could not create Excel workbook.');
		}

		// add relations
		$oZip->addFromString('_rels' . DIRECTORY_SEPARATOR . '.rels', $this->documentRelations());

		// add content types
		$oZip->addFromString('[Content_Types].xml', $this->documentContentTypes());

		// add app.xml
		$oZip->addFromString('docProps' . DIRECTORY_SEPARATOR . 'app.xml', $this->workbookAppXml());

		// add core.xml
		$oZip->addFromString('docProps' . DIRECTORY_SEPARATOR . 'core.xml', $this->workbookCoreXml());

		// add workbook relations
		$oZip->addFromString('xl' . DIRECTORY_SEPARATOR . '_rels' . DIRECTORY_SEPARATOR . 'workbook.xml.rels', $this->workbookRelations());

		// add Calc chain if exists
		if (!empty($this->sCalcChain)) {
			$oZip->addFromString('xl' . DIRECTORY_SEPARATOR . 'calcChain.xml', $this->calcChain());
		}

		// add shared strings
		if (!empty($this->aSharedStrings)) {
			$oZip->addFromString('xl' . DIRECTORY_SEPARATOR . 'sharedStrings.xml', $this->sharedStrings());
		} else {
			echo 'no shared strings';
			print_r($this->aSharedStrings);
		}

		// add styles
		$oZip->addFromString('xl' . DIRECTORY_SEPARATOR . 'styles.xml', $this->workbookStyles());

		// add workbook theme.xml
		$oZip->addFromString('xl' . DIRECTORY_SEPARATOR . 'theme' . DIRECTORY_SEPARATOR . 'theme1.xml', $this->workbookTheme());

		// add workbook information
		$oZip->addFromString('xl' . DIRECTORY_SEPARATOR . 'workbook.xml', $this->sWorkbook);

		foreach($this->aSheetData as $sSheetId => $sDom) {
			$oZip->addFromString('xl' . DIRECTORY_SEPARATOR . 'worksheets' . DIRECTORY_SEPARATOR . $sSheetId . '.xml', $sDom);
		}

		$oZip->close();
	}

	/**
	 * create workbook relations xml string
	 *
	 * @return string
	 */
	private function documentRelations() {
		return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
			. '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    		. '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
    		. '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
    		. '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
			. '</Relationships>'
		;
	}

	/**
	 * create [Content_Types.xml] string
	 *
	 * @return string
	 */
	private function documentContentTypes() {
		$sXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
			. '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
			. '<Default Extension="xml" ContentType="application/xml"/>'
			. '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
//			. '<Default Extension="jpeg" ContentType="image/jpeg"/>'
			. '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
			. '<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
			. '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'

			// add shared strings if exists
			. ((!empty($this->aSharedStrings)) ? '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' : '')

			// add calc chain if exists
			. ((!empty($this->sCalcChain)) ? '<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>' : '')

			. '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
			. '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'

			// add sheets
			. '%s'
			. '</Types>'
		;

		// add defined sheet files
		$sSheets = '';

		// template for sheet ovverride
		$sSheetTemplate = '<Override PartName="/xl/worksheet/%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';

		// iterate dom sheets
		foreach($this->aSheets as $oSheet) {
			/** @var Sheet $oSheet */

			// add filename by sheet identifier
			$sSheets .= sprintf($sSheetTemplate, $oSheet->getIdentifier());
		}

		// return generated xml string
		return sprintf($sXml, $sSheets);
	}

	/**
	 * create app.xml
	 *
	 * @return string
	 */
	private function workbookAppXml() {

		// Create Xml
		$sXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
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
			. '<vt:i4>%1$d</vt:i4>' // replaces %d with number of sheets
			. '</vt:variant>'
			. '</vt:vector>'
			. '</HeadingPairs>'
			. '<TitlesOfParts>'
			. '<vt:vector size="%1$d" baseType="lpstr">%2$s</vt:vector>' // replaces %s with sheet names
			. '	</TitlesOfParts>'
			. '	<Company>Excellence</Company>'
			. '	<LinksUpToDate>false</LinksUpToDate>'
			. '	<SharedDoc>false</SharedDoc>'
			. '	<HyperlinksChanged>false</HyperlinksChanged>'
			. '	<AppVersion>14.0300</AppVersion>'
			. '</Properties>'
		;

		$sSheets = '';
		$sSheetTemplate = '<vt:lpstr>%s</vt:lpstr>';

		// iterate dom sheets
		foreach($this->aSheets as $oSheet) {

			// add filename by sheet identifier
			$sSheets .= sprintf($sSheetTemplate, $oSheet->getName());
		}

		return sprintf($sXml, $this->iSheets, $sSheets);
	}

	/**
	 * create core.xml including metadata
	 * 
	 * @return string
	 */
	private function workbookCoreXml() {
		return  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
			. '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
			// @ todo insert creators name - metata delegate
			. '<dc:creator>creator name</dc:creator>'
    		. '<cp:lastModifiedBy>Name of last modified</cp:lastModifiedBy>'
    		. '<dcterms:created xsi:type="dcterms:W3CDTF">' . date('Y-m-d\TH:i:s\Z') . '</dcterms:created>'
    		. '<dcterms:modified xsi:type="dcterms:W3CDTF">' . date('Y-m-d\TH:i:s\Z') . '</dcterms:modified>'
			. '</cp:coreProperties>'
		;
	}

	/**
	 * create workbook relations as string because it is faster than creating dom documents
	 *
	 * @return string
	 */
	private function workbookRelations() {

		// Xml
		$sXml = '<?xml version="1.0"?>'
			. '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
				. '%s' // workbook sheets
				. '<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>'
				. '<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
			. '</Relationships>'
		;

		$sSheets = '';
		$sSheetTemplate = '<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/%s.xml"/>';
		$iId = 1;
		// iterate dom sheets
		foreach($this->aSheets as $oSheet) {
			/** @var Sheet $oSheet */

			// add filename by sheet identifier
			$sSheets .= sprintf($sSheetTemplate, $iId, $oSheet->getIdentifier());

			// increase id
			$iId ++;
		}

		// add calc chain
		if (!empty($this->sCalcChain)) {
			$sSheets .= sprintf('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>', $iId);
			$iId++;
		}

		// add Shared Strings
		if (!empty($this->aSharedStrings)) {
			$sSheets .= sprintf('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>', $iId+2);
		}

		// return generated XML
		return sprintf($sXml, $sSheets, $iId, $iId + 1);
	}

	/**
	 * create workbook styles
	 *
	 * @return string
	 */
	private function workbookStyles() {
		return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
			. '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">'
				. '<fonts count="1" x14ac:knownFonts="1">'
					. '<font>'
						. '<sz val="12"/>'
						. '<color theme="1"/>'
						. '<name val="Calibri"/>'
						. '<family val="2"/>'
						. '<scheme val="minor"/>'
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
					. '<border>'
						. '<left/>'
						. '<right/>'
						. '<top/>'
						. '<bottom/>'
						. '<diagonal/>'
					. '</border>'
				. '</borders>'
				. '<cellStyleXfs count="1">'
					. '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>'
				. '</cellStyleXfs>'
				. '<cellXfs count="1">'
					. '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
				. '</cellXfs>'
				. '<cellStyles count="1">'
					. '<cellStyle name="Standard" xfId="0" builtinId="0"/>'
				. '</cellStyles>'
				. '<dxfs count="0"/>'
				. '<tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleMedium4"/>'
			. '</styleSheet>'
		;
	}

	/**
	 * create workbook theme - standard Excel theme
	 *
	 * @return string
	 */
	private function workbookTheme() {
		// standard theme
		return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office-Design"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Cambria"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="宋体"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="宋体"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="1"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="100000"/><a:shade val="100000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="50000"/><a:shade val="100000"/><a:satMod val="350000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"><a:shade val="95000"/><a:satMod val="105000"/></a:schemeClr></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="38000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst><a:scene3d><a:camera prst="orthographicFront"><a:rot lat="0" lon="0" rev="0"/></a:camera><a:lightRig rig="threePt" dir="t"><a:rot lat="0" lon="0" rev="1200000"/></a:lightRig></a:scene3d><a:sp3d><a:bevelT w="63500" h="25400"/></a:sp3d></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="40000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="40000"><a:schemeClr val="phClr"><a:tint val="45000"/><a:shade val="99000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="20000"/><a:satMod val="255000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="-80000" r="50000" b="180000"/></a:path></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="80000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="30000"/><a:satMod val="200000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults><a:spDef><a:spPr/><a:bodyPr/><a:lstStyle/><a:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="3"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="2"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></a:style></a:spDef><a:lnDef><a:spPr/><a:bodyPr/><a:lstStyle/><a:style><a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="1"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></a:style></a:lnDef></a:objectDefaults><a:extraClrSchemeLst/></a:theme>';
	}

	/**
	 * Generate shared strings xml document.
	 *
	 * @return string
	 */
	private function sharedStrings() {
		$sXml = '<?xml version="1.0" encoding="utf-8"?>'
			. '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
				. 'count="' . count($this->aSharedStrings) . '" uniqueCount="' . count($this->aSharedStrings) . '">'
		;

		foreach($this->aSharedStrings as $sString => $iIndex) {
			$sXml .= '<si><t>' . $sString . '</t></si>';
		}

		return $sXml . '</sst>';
	}

	/**
	 * Generate calc chain document
	 *
	 * @return string
	 */
	private function calcChain() {
		return '<?xml version="1.0" encoding="utf-8"?>'
			. '<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
			. $this->sCalcChain
			. '</calcChain>'
		;

	}
}