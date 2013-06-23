<?php
/**
 * @author        Ronald Marske <scyks@ceow.de>
 * @filesource    Workbook.php
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

#pragma mark - constants

	const TYPE_STRING    = 1;
	const TYPE_NUMBER    = 2;
	const TYPE_FUNCTION  = 4;

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
	 * @var \DomDocument|null
	 */
	private $oSheets;

	/**
	 * contains an array [sheet identifier] => DomDocument including sheet data.
	 *
	 * @var array
	 */
	private $aSheetData;

	/**
	 * contains calc chain document
	 *
	 * @var null|\DomDocument
	 */
	private $oCalcChain = null;

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
		$iSheets = (int) $this->oDelegate->numberOfSheetsInWorkbook($this);

		// make sure that there is minimum one sheet
		if (0 >= $iSheets) {
			throw new \LogicException('WorkbookDelegate::numberOfSheetsInWorkbook have to return an integer bigger than zero.');
		}

		// create Sheet Xml Data
		$this->createSheetXml($iSheets);

		return $this;

	}

	/**
	 * creates DomDocument for Excel sheets that would be used to create an XML
	 * by using XSLT Stylesheet. Of course it is possible to create real OpenXML
	 * Document, but with an XSLT it is easier to handle namespaces, format
	 * changes and further development s of this library.
	 *
	 * @param int $iSheets
	 *
	 * @throws \LogicException
	 */
	private function createSheetXml($iSheets) {

		$this->oSheets = new \DOMDocument('1.0', 'utf-8');

		// sheets node
		$oSheets = $this->oSheets->createElement('sheets');
		$oSheets->setAttribute('count', $iSheets);

		// add sheets node to document
		$this->oSheets->appendChild($oSheets);

		// iterate sheets
		for($iSheet = 0; $iSheet < $iSheets; $iSheet++) {

			/** @var Sheet $oSheet */
			$oSheet = $this->oDelegate->getSheetForWorkBook($this, $iSheet);

			// make sure that oSheet is instanceof sheet
			if (!$oSheet instanceof Sheet) {
				throw new \LogicException(sprintf('WorkbookDelegate::getSheetForWorkBook have to return an instance of \Excellence\Sheet, "%s" given.', gettype($oSheet)));
			}

			// create sheet node
			$oXmlSheet = $this->oSheets->createElement('sheet', $oSheet->getName());
			$oXmlSheet->setAttribute('id', $oSheet->getIdentifier());

			// append sheet to sheets node
			$oSheets->appendChild($oXmlSheet);

			// create sheet data by sheet
			$this->createSheetDataXml($oSheet, $iSheet + 1);
		}

		// unset variables
		unset($oSheets, $oXmlSheet, $iSheet, $iSheets);

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
		$oDataSource = $this->oDelegate->dataSourceForWorkbookAndSheet($this, $oSheet);

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


		$oDom = new \DOMDocument('1.0', 'utf-8');
		$oSheets = $oDom->createElement('sheets');
		$oSheets->setAttribute('id', $oSheet->getIdentifier());
		$oSheets->setAttribute('rows', $iRows);
		$oSheets->setAttribute('columns', $iColumns);

		$oDom->appendChild($oSheets);

		$sDimensionFrom = null;
		$sDimensionTo = null;

		// sheet loop
		for($iRow = 0; $iRow < $iRows; $iRow++) {

			// create row
			$oRow = $oDom->createElement('row');
			$oRow->setAttribute('id', $iRow + 1);

			// append row to sheets
			$oSheets->appendChild($oRow);

			for($iColumn = 0; $iColumn < $iColumns; $iColumn++) {

				// get value for cell
				$value = $oDataSource->valueForRowAndColumn($this, $oSheet, $iRow, $iColumn);

				if (null === $value) continue;

				// value type
				$iType = gettype($value);

				// make sure we get an allowed data type
				if (!in_array($iType, array('string', 'integer', 'float', 'double'))) {
					throw new \LogicException(sprintf('DataDelegate::valueForRowAndColumn have to return a string, float, double or int value, "%s" given.', $iType));
				}

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

				// determine value type
				if ('string' == $iType && '=' == substr($value, 0, 1)) {
					$iType = self::TYPE_FUNCTION;
					$value = substr($value, 1);

					$this->addColumnToCalcChain($sCord, $iSheet);

				} elseif('string' == $iType) {
					$iType = self::TYPE_STRING;
				} else {
					$iType = self::TYPE_NUMBER;
				}

				// create column
				$oColumn = $oDom->createElement('column', $value);
				$oColumn->setAttribute('id', $sCord);
				$oColumn->setAttribute('type', $iType);

				// add column to row
				$oRow->appendChild($oColumn);
			}
		}

		$oSheets->setAttribute('dimension', $sDimensionFrom . ':' . $sDimensionTo);
		$oSheets->setAttribute('dimension', $sDimensionFrom . ':' . $sDimensionTo);

		$this->aSheetData[$oSheet->getIdentifier()] = $oDom;
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
	 * this method will create one.
	 *
	 * @param string $sCord
	 * @param int $iSheet
	 */
	private function addColumnToCalcChain($sCord, $iSheet) {

		if (null == $this->oCalcChain) {
			$this->oCalcChain = new \DOMDocument('1.0', 'utf-8');
			$oCalcChain = $this->oCalcChain->createElementNS('http://schemas.openxmlformats.org/spreadsheetml/2006/main', 'calcChain', '');
			$this->oCalcChain->appendChild($oCalcChain);
		} else {
			$oCalcChain = $this->oCalcChain->getElementsByTagName('calcChain')->item(0);
		}

		$oColumn = $this->oCalcChain->createElement('c');
		$oColumn->setAttribute('r', $sCord);
		$oColumn->setAttribute('i', $iSheet);

		$oCalcChain->appendChild($oColumn);


	}

#pragma mark - saving and creating officeOpenDocument files

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
		if ($this->oCalcChain instanceof \DomDocument) {
			$oZip->addFromString('xl' . DIRECTORY_SEPARATOR . 'calcChain.xml', $this->oCalcChain->saveXML());
		}
		// add styles
		$oZip->addFromString('xl' . DIRECTORY_SEPARATOR . 'styles.xml', $this->workbookStyles());

		// add workbook theme.xml
		$oZip->addFromString('xl' . DIRECTORY_SEPARATOR . 'theme' . DIRECTORY_SEPARATOR . 'theme1.xml', $this->workbookTheme());

		// add workbook information
		$oZip->addFromString('xl' . DIRECTORY_SEPARATOR . 'workbook.xml', $this->workbook());

		$this->workbookSheets($oZip);

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
//			. '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'

			// add calchain if exists
			. ((null !== $this->oCalcChain) ? '<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>' : '')
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
		foreach($this->oSheets->getElementsByTagName('sheet') as $oSheet) {
			/** @var \DomElement $oSheet */

			// add filename by sheet identifier
			$sSheets .= sprintf($sSheetTemplate, $oSheet->getAttribute('id'));
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
							. '<vt:i4>%d</vt:i4>' // replaces %d with number of sheets
						. '</vt:variant>'
					. '</vt:vector>'
				. '</HeadingPairs>'
				. '<TitlesOfParts>'
					. '<vt:vector size="2" baseType="lpstr">%s</vt:vector>' // replaces %s with sheet names
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
		$iSheets = 0;
		// iterate dom sheets
		foreach($this->oSheets->getElementsByTagName('sheet') as $oSheet) {
			/** @var \DomElement $oSheet */

			// add filename by sheet identifier
			$sSheets .= sprintf($sSheetTemplate, $oSheet->nodeValue);

			// @todo: save number of sheets by asking delegate in method createSheetXml
			// increase sheet count
			$iSheets ++;
		}

		return sprintf($sXml, $iSheets, $sSheets);
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
	 * create workbook relations
	 *
	 * @return string
	 */
	private function workbookRelations() {
		$sXml = '<?xml version="1.0"?>'
			. '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
				. '%s' // workbook sheets
//				. '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
//				. '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>'
//				. '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
//				. '<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>'
				. '<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>'
				. '<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
			. '</Relationships>'
		;

		$sSheets = '';
		$sSheetTemplate = '<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/%s.xml"/>';
		$iSheets = 1;
		// iterate dom sheets
		foreach($this->oSheets->getElementsByTagName('sheet') as $oSheet) {
			/** @var \DomElement $oSheet */

			// add filename by sheet identifier
			$sSheets .= sprintf($sSheetTemplate, $iSheets, $oSheet->getAttribute('id'));

			// @todo: save number of sheets by asking delegate in method createSheetXml
			// increase sheet count
			$iSheets ++;
		}

		if ($this->oCalcChain instanceof \DOMDocument) {
			$sSheets .= sprintf('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>', $iSheets);
			$iSheets++;
		}

		return sprintf($sXml, $sSheets, $iSheets, $iSheets + 1);
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
	 * create workbook informations
	 *
	 * @return string
	 */
	private function workbook() {
		$sXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
			. '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    			. '<fileVersion appName="xl" lastEdited="5" lowestEdited="5" rupBuild="23206"/>'
    			. '<workbookPr showInkAnnotation="0" autoCompressPictures="0"/>'
    			. '<bookViews>'
        			. '<workbookView xWindow="0" yWindow="0" windowWidth="25600" windowHeight="14460" tabRatio="500"/>'
    			. '</bookViews>'
    			. '<sheets>'
					. '%s' // %s is replaced by sheet information
				. '</sheets>'
    			. '<calcPr calcId="140000" concurrentCalc="0"/>'
    			. '<extLst>'
        			. '<ext xmlns:mx="http://schemas.microsoft.com/office/mac/excel/2008/main" uri="{7523E5D3-25F3-A5E0-1632-64F254C22452}">'
            			. '<mx:ArchID Flags="2"/>'
        			. '</ext>'
    			. '</extLst>'
			. '</workbook>'
		;

		$sSheets = '';
		$sSheetTemplate = '<sheet name="%1$s" sheetId="%2$d" r:id="rId%2$d"/>';
		$iSheets = 1;
		// iterate dom sheets
		foreach($this->oSheets->getElementsByTagName('sheet') as $oSheet) {
			/** @var \DomElement $oSheet */

			// add filename by sheet identifier
			$sSheets .= sprintf($sSheetTemplate, $oSheet->nodeValue, $iSheets);

			// @todo: save number of sheets by asking delegate in method createSheetXml
			// increase sheet count
			$iSheets ++;
		}

		return sprintf($sXml, $sSheets);
	}

	/**
	 * generate sheet data
	 * @param \ZipArchive $oZip
	 */
	private function workbookSheets(\ZipArchive $oZip) {

		$oXslt = new \XSLTProcessor();

		$oXsl = new \DOMDocument();
		$oXsl->load(dirname(__FILE__) . DIRECTORY_SEPARATOR . 'workbook.xsl', LIBXML_NOCDATA);

		$oXslt->importStylesheet($oXsl);

		foreach($this->aSheetData as $sSheetId => $oDom) {

			#echo $oXslt->transformToXml($oDom);
			$oZip->addFromString(
				'xl' . DIRECTORY_SEPARATOR . 'worksheets' . DIRECTORY_SEPARATOR . $sSheetId . '.xml',
				$oXslt->transformToXml($oDom)
			);
		}

	}
}