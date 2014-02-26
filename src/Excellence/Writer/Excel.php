<?php
/**
 * @author		  Ronald Marske <r.marske@secu-ring.de>
 * @filesource	  src/Excellence/Writer/Excel.php
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

namespace Excellence\Writer;

use Excellence\Sheet;
use Excellence\Style;
use Excellence\Delegates\DataDelegate;
use Excellence\Delegates\StylableDelegate;
use Excellence\Delegates\MergeableDelegate;
use Excellence\Delegates\LinkableDelegate;


/**
 * Excel document writer class
 *
 * @package Excellence\Writer
 */
class Excel extends AbstractWriter {

#pragma mark - properties

	/**
	 * contains number of sheets in document
	 * @var int
	 */
	protected $iSheets = null;

	/**
	 * contains dom document including sheets
	 *
	 * @var array
	 */
	protected $aSheets = array();

	/**
	 * contains an array [sheet identifier] => DomDocument including sheet data.
	 *
	 * @var array
	 */
	protected $aSheetData;

	/**
	 * contains array with shared strings
	 * string => id
	 * @var array
	 */
	private $aSharedStrings = array();

	/**
	 * contains array with hyperlinks
	 * hyperlink => id
	 * @var array
	 */
	private $aHyperlinks = array();

	/**
	 * array containing styles for workbook
	 * @var array
	 */
	private $aStyles = array();

	/**
	 * Style References
	 * @var array
	 */
	private $aStyleRefs = array();

#pragma mark - sharedStrings

	/**
	 * add a string to shared strings and returns index of this value
	 * @param string $value
	 *
	 * @return int
	 */
	protected function addValueToSharedStrings($value) {

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

#pragma mark - hyperlinks

	/**
	 * add a hyperlink to link list and returns index of this hyperlink
	 *
	 * @param Sheet $oSheet
	 * @param string $value
	 *
	 * @return int
	 */
	protected function addHyperlink(Sheet &$oSheet, $value) {

		if (!array_key_exists($oSheet->getIdentifier(), $this->aHyperlinks)) {
			$this->aHyperlinks[$oSheet->getIdentifier()] = array();
		}

		// check if value currently exists and return index
		if (array_key_exists($value, $this->aHyperlinks[$oSheet->getIdentifier()])) {
			return $this->aHyperlinks[$oSheet->getIdentifier()][$value];
		}

		// get current count
		$iNum = count($this->aHyperlinks[$oSheet->getIdentifier()]) + 1;

		// add value to shared strings table
		$this->aHyperlinks[$oSheet->getIdentifier()][$value] = $iNum;

		// return current index;
		return $iNum;

	}

	/**
	 * will check if there are hyperlinks and returns true, when there is
	 * minimum one hyperlink. Otherwise this method returns false.
	 *
	 * @param Sheet $oSheet
	 * @return bool
	 */
	protected function hasHyperlinks(Sheet &$oSheet) {
		return (isset($this->aHyperlinks[$oSheet->getIdentifier()]) && !empty($this->aHyperlinks[$oSheet->getIdentifier()]));
	}

#pragma mark - sheet handling

	/**
	 * get number of sheets from workbook and stores it for performance
	 * @return int
	 */
	public function getNumberOfSheets() {

		if (null == $this->iSheets)
			$this->iSheets = (int) $this->getDelegate()->numberOfSheetsInWorkbook($this->getWorkbook());

		return $this->iSheets;
	}

	/**
	 * make sure that there is minimum one sheet in workbook
	 * @return $this
	 * @throws \LogicException
	 */
	protected function checkNumberOfSheetsInWorkbook() {

		if (0 >= $this->getNumberOfSheets()) {
			throw new \LogicException('WorkbookDelegate::numberOfSheetsInWorkbook have to return an integer bigger than zero.');
		}

		return $this;
	}

	/**
	 * this method will generate xml sources for sheet information
	 *
	 * @throws \LogicException
	 */
	protected function createSheetXml() {

		$sWorkbook = '<?xml version="1.0" encoding="UTF-8"?>' . "\n"
			. '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><fileVersion appName="xl" lastEdited="5" lowestEdited="5" rupBuild="23206"/><workbookPr showInkAnnotation="0" autoCompressPictures="0"/>'
			. '<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="25600" windowHeight="14460" tabRatio="500"/></bookViews>'
			. '<sheets>';

		// iterate sheets
		for($iSheet = 0; $iSheet < $this->getNumberOfSheets(); $iSheet++) {

			/** @var Sheet $oSheet */
			$oSheet = $this->getDelegate()->getSheetForWorkBook($this->getWorkbook(), $iSheet);

			// make sure that oSheet is instanceof sheet
			if (!$oSheet instanceof Sheet) {
				throw new \LogicException(sprintf('WorkbookDelegate::getSheetForWorkBook have to return an instance of \Excellence\Sheet, "%s" given.', gettype($oSheet)));
			}

			$this->aSheets[] = $oSheet;

			$sWorkbook .= '<sheet name="' . (($oSheet->hasName()) ? $oSheet->getName() : 'Sheet ' . ($iSheet + 1)) . '" sheetId="' . ($iSheet + 1) . '" r:id="rId' . ($iSheet + 1) . '"/>';

		}

		$sWorkbook .= '</sheets><calcPr calcId="140000" concurrentCalc="0"/><extLst><ext xmlns:mx="http://schemas.microsoft.com/office/mac/excel/2008/main" uri="{7523E5D3-25F3-A5E0-1632-64F254C22452}"><mx:ArchID Flags="2"/></ext></extLst></workbook>';

		return $sWorkbook;

	}

	/**
	 * generate values for given sheet.
	 *
	 * @param Sheet $oSheet
	 * @param int $iSheet
	 *
	 * @throws \LogicException
	 */
	protected function createSheetDataXml(Sheet $oSheet, $iSheet) {

		/** @var DataDelegate $oDataSource */
		$oDataSource = $this->getDelegate()->dataSourceForWorkbookAndSheet($this->getWorkbook(), $oSheet);

		if (!$oDataSource instanceof DataDelegate) {
			throw new \LogicException('WorkbookDelegate::dataSourceForWorkbookAndSheet have to return an instance of \Excellence\Delegates\DataSource, "NULL" given.');
		}

		// get rows
		$iRows = (int) $oDataSource->numberOfRowsInSheet($this->getWorkbook(), $oSheet);

		// throw exception if retrieving other value than positive inter bigger than zero
		if (0 >= $iRows) {
			throw new \LogicException('DataDelegate::numberOfRowsInSheet have to return an integer value bigger than zero.');
		}

		// get number of columns
		$iColumns = (int) $oDataSource->numberOfColumnsInSheet($this->getWorkbook(), $oSheet);

		if (0 >= $iColumns) {
			throw new \LogicException('DataDelegate::numberOfColumnsInSheet have to return an integer value bigger than zero.');
		}

		// get default styles
		if ($oDataSource instanceof StylableDelegate) {

			// set default styles
			$oStandardStyle = $oDataSource->getStandardStyle($this->getWorkbook(), $oSheet);

			// if style is instance of Style, set it
			if ($oStandardStyle instanceof Style) {
				$this->getWorkbook()->getStandardStyles($oStandardStyle);
			}
		}

		// create workbook xml
		$this->aSheetData[$oSheet->getIdentifier()] = '<?xml version="1.0" encoding="UTF-8"?>' . "\n"
			. '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">'
			// formally at this position is a dimension tag, but Excel don't need it <dimension ref="A1:B2"/>
			. '<sheetViews>'
		;

		if ($oSheet->isFirstRowFixed()) {
			$this->aSheetData[$oSheet->getIdentifier()] .= '<sheetView tabSelected="1" workbookViewId="0">'
				. '<pane ySplit="1" topLeftCell="A2" state="frozen"/>'
				. '</sheetView>'
			;
		} else {
			$this->aSheetData[$oSheet->getIdentifier()] .= '<sheetView tabSelected="1" workbookViewId="0"/>';
		}

		$this->aSheetData[$oSheet->getIdentifier()] .= '</sheetViews>'
			. '%s'
			. '<sheetData>'
		;

		$sMerge = '';
		$aColStyles = array();
		$sHyperlinks = '';

		// sheet loop
		for($iRow = 0; $iRow < $iRows; $iRow++) {


			$sRowStyles = '';
			$sRow = '';
			// create row
			$sRow .= "\n".'<row r="' . ($iRow + 1) . '"%s>';

			for($iColumn = 0; $iColumn < $iColumns; $iColumn++) {

				// get value for cell
				$value = $oDataSource->valueForRowAndColumn($this->getWorkbook(), $oSheet, $iRow, $iColumn);

				// Excel coordinate
				$sCord = $this->getWorkbook()->getCoordinatesByColumnAndRow($iColumn, $iRow + 1);

				// style definition for cell
				$sStyle = null;

				// return if value is empty
				if (null === $value) continue;

				// value type
				$iType = gettype($value);

				// make sure we get an allowed data type
				if (!in_array($iType, array('string', 'integer', 'float', 'double', 'boolean'))) {
					throw new \LogicException(sprintf('DataDelegate::valueForRowAndColumn have to return a string, float, double or int value, "%s" given.', $iType));
				}

				// check if data source implements mergeable delegate
				if ($oDataSource instanceof MergeableDelegate) {

					// get merge range
					$sMergeRange = $this->getDelegate()->mergeByColumnAndRow($this->getWorkbook(), $iColumn, $iRow);

					// continue only if mergeRange is a string
					if (true == is_string($sMergeRange)) {

						// make sure this is upper cased
						$sMergeRange = strtoupper($sMergeRange);

						// check if merge range equals correct format
						if (!preg_match('/^[A-Z]+[0-9]+:[A-Z]+[0-9]$/', $sMergeRange)) {
							throw new \InvalidArgumentException('MergeableDelegate::mergeByColumnAndRow have to return a merge range as string in format "A1:A2"');
						}

						$sMerge .= '<mergeCell ref="' . $sMergeRange . '"/>';
					}
				}


				// check if data source is instance of styleable delegate
				if ($oDataSource instanceof StylableDelegate) {
					$oStyle = $oDataSource->getStyleForColumnAndRow($this->getWorkbook(), $oSheet, $iColumn, $iRow);

					if ($oStyle instanceof Style) {
						if (!array_key_exists($oStyle->getId(), $this->aStyles)) {
							$this->aStyles[$oStyle->getId()] = $oStyle;
							$this->aStyleRefs[$oStyle->getId()] = count($this->aStyles);
						}

						$sStyle = ' s="' . $this->aStyleRefs[$oStyle->getId()] . '"';

						// height for columns
						if ($oStyle->hasHeight() && '' == $sRowStyles) {
							$sRowStyles .= ' spans="1:8" customFormat="1" ht="'.$oStyle->getHeight().'" customHeight="1"';
						}

						// width for columns
						if ($oStyle->hasWidth() && !array_key_exists($iColumn, $aColStyles)) {
							$aColStyles[$iColumn] = '<col min="'.($iColumn+1).'" max="'.($iColumn+1).'" width="'.$oStyle->getWidth().'" customWidth="1"/>';
						}
					}

				}

				// Hyperlinks
				if ($oDataSource instanceof LinkableDelegate && $oDataSource->hasLinkForColumnAndRow($this->getWorkbook(), $oSheet, $iRow, $iColumn)) {

					$iLinkId = $this->addHyperlink($oSheet, $oDataSource->getLinkForColumnAndRow($this->getWorkbook(), $oSheet, $iRow, $iColumn));
					$sHyperlinks .= '<hyperlink ref="' . $sCord . '" r:id="rId' . $iLinkId . '" />';

				}

				// function or formula
				if ('string' == $iType && '=' == mb_substr($value, 0, 1, 'utf-8')) {

					// add value to column
					$sRow .= '<c r="' . $sCord . '"' . $sStyle . '><f>' . substr($value, 1) . '</f></c>';

					// string
				} elseif('string' == $iType) {
					$iNum = $this->addValueToSharedStrings($value);

					// add value to column
					$sRow .= '<c r="' . $sCord . '" t="s"' . $sStyle . '><v>' . $iNum . '</v></c>';

					// boolean
				} elseif ('boolean' == $iType) {

					// add value to column
					$sRow .= '<c r="' . $sCord . '" t="b"' . $sStyle . '><v>' . (int) $value . '</v></c>';

					// number
				} else {

					// add value to column
					$sRow .= '<c r="' . $sCord . '" t="n"' . $sStyle . '><v>' . $value . '</v></c>';
				}

			}

			$sRow .= '</row>';

			// replace row styles
			$sRow = sprintf($sRow, $sRowStyles);

			$this->aSheetData[$oSheet->getIdentifier()] .= $sRow;
		}

		// set col styles
		if (!empty($aColStyles)) {
			$this->aSheetData[$oSheet->getIdentifier()] = sprintf($this->aSheetData[$oSheet->getIdentifier()], '<cols>'.implode('', $aColStyles).'</cols>');
		} else {
			$this->aSheetData[$oSheet->getIdentifier()] = sprintf($this->aSheetData[$oSheet->getIdentifier()], '');
		}

		$this->aSheetData[$oSheet->getIdentifier()] .= '</sheetData>';

		// add hyperlinks
		if (0 < strlen($sHyperlinks)) {
			$this->aSheetData[$oSheet->getIdentifier()] .= '<hyperlinks>' . $sHyperlinks . '</hyperlinks>';
			unset($sHyperlinks);
		}

		// merged cells
		if (0 < strlen($sMerge)) {
			$this->aSheetData[$oSheet->getIdentifier()] .= '<mergeCells>' . $sMerge . '</mergeCells>';
			unset($sMerge);unset($sMerge);
		}

		$this->aSheetData[$oSheet->getIdentifier()] .= '</worksheet>';

	}

#pragma mark - document Relations

	/**
	 * create workbook relations xml string
	 *
	 * @return string
	 */
	protected  function documentRelationsXml() {
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
	protected function documentContentTypesXml() {
		$sXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
			. '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
			. '<Default Extension="xml" ContentType="application/xml"/>'
			. '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
			. '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
			. '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'

			// add shared strings if exists
			. ((!empty($this->aSharedStrings)) ? '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' : '')

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
	protected function workbookAppXml() {

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
			. '</TitlesOfParts>'
			. '<Company>Excellence</Company>'
			. '<LinksUpToDate>false</LinksUpToDate>'
			. '<SharedDoc>false</SharedDoc>'
			. '<HyperlinksChanged>false</HyperlinksChanged>'
			. '<AppVersion>14.0300</AppVersion>'
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
	protected function workbookCoreXml() {
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
	protected function workbookRelationsXml() {

		// Xml
		$sXml = '<?xml version="1.0"?>'
			. '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
			. '%s' // workbook sheets
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

		// add Shared Strings
		if (!empty($this->aSharedStrings)) {
			$sSheets .= sprintf('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>', $iId+2);
		}

		// return generated XML
		return sprintf($sXml, $sSheets, $iId);
	}

	/**
	 * create workbook styles
	 *
	 * @return string
	 */
	protected function workbookStylesXML() {

		$aFonts = array();
		$aFills = array(
			'none' => '<fill><patternFill patternType="none"/></fill>',
			'gray125' => '<fill><patternFill patternType="gray125"/></fill>',
		);
		$aBorders = array(
			'<border/>' => 0,
		);
		$aStyles = array(
			'<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" shrinkToFit="true" wrapText="true"/>',
		);


		$getFontName = function(Style $oStyle) {
			return $oStyle->getFont()
				. '_' . $oStyle->getFontSize()
				. '_' . $oStyle->getColor()
				. ((true == $oStyle->isBold()) ? '_bold' : '')
				. ((true == $oStyle->isItalic()) ? '_italic' : '')
				. ((true == $oStyle->hasUnderline()) ? '_underline' : '')
			;
		};

		$addFont = function(Style $oStyle) use(&$aFonts, $getFontName) {
			$sFontName = $getFontName($oStyle);

			if (!array_key_exists($sFontName, $aFonts)) {

				$aFonts[$sFontName] = '<font>';

				if ($oStyle->hasFontSize()) {
					$aFonts[$sFontName] .= '<sz val="' . $oStyle->getFontSize() . '"/>';
				}

				if ($oStyle->hasFont()) {
					$aFonts[$sFontName] .= '<name val="' . $oStyle->getFont() . '"/>';

				}

				if ($oStyle->hasColor()) {
					$aFonts[$sFontName] .= '<color rgb="FF' . $oStyle->getColor() . '"/>';
				}

				if ($oStyle->isBold()) {
					$aFonts[$sFontName] .= '<b/>';
				}

				if ($oStyle->isItalic()) {
					$aFonts[$sFontName] .= '<i/>';
				}

				if ($oStyle->hasUnderline()) {
					$aFonts[$sFontName] .= '<u/>';
				}

				$aFonts[$sFontName] .= '</font>';

				if ('<font></font>' == $aFonts[$sFontName]) {
					unset($aFonts[$sFontName]);
					return 0;
				}
			}

			return array_search($sFontName, array_keys($aFonts));


		};

		$addFill = function(Style $oStyle) use(&$aFills) {
			if ($oStyle->hasBackgroundColor()) {
				if (!array_key_exists($oStyle->getBackgroundColor(), $aFills)) {
					$aFills[$oStyle->getBackgroundColor()] = '<fill><patternFill patternType="solid"><fgColor rgb="FF'.$oStyle->getBackgroundColor().'"/></patternFill></fill>';
				}

				return array_search($oStyle->getBackgroundColor(), array_keys($aFills));
			}
			return 0;
		};

		$addBorder = function(Style $oStyle) use(&$aBorders) {
			if ($oStyle->hasBorder()) {

				$sBorder = '<border>';
				foreach($oStyle->getBorder() as $iAlignment => $aStyle) {

					if ($iAlignment == Style::BORDER_ALIGN_LEFT) {
						$sBorder .= '<left style="'.$aStyle['style'].'"><color rgb="FF'.$aStyle['color'].'"/></left>';
					}

					if ($iAlignment == Style::BORDER_ALIGN_RIGHT) {
						$sBorder .= '<right style="'.$aStyle['style'].'"><color rgb="FF'.$aStyle['color'].'"/></right>';
					}

					if ($iAlignment == Style::BORDER_ALIGN_TOP) {
						$sBorder .= '<top style="'.$aStyle['style'].'"><color rgb="FF'.$aStyle['color'].'"/></top>';
					}

					if ($iAlignment == Style::BORDER_ALIGN_BOTTOM) {
						$sBorder .= '<bottom style="'.$aStyle['style'].'"><color rgb="FF'.$aStyle['color'].'"/></bottom>';
					}
				}


				$sBorder .= '</border>';

				if (!array_key_exists($sBorder, $aBorders)) {
					$aBorders[$sBorder] = count($aBorders);
				}

				return $aBorders[$sBorder];
			}
			return count($aBorders) - 1;
		};

		if ($this->getWorkbook()->getStandardStyles() instanceof Style) {
			$addFont($this->getWorkbook()->getStandardStyles());
		}

		/** @var Style $oStyle */
		foreach($this->aStyles as $oStyle) {

			$sStyle = '<xf xfId="0" numFmtId="0" fontId="'.$addFont($oStyle).'" fillId="'.$addFill($oStyle).'" borderId="'.$addBorder($oStyle).'" shrinkToFit="true" wrapText="true"';

			if ($oStyle->hasHorizontalAlignment() || $oStyle->hasVerticalAlignment()) {
				$sStyle .= '>';

				$sStyle .= '<alignment horizontal="'.$oStyle->getHorizontalAlignment().'" vertical="'.$oStyle->getVerticalAlignment().'"/>';
				$sStyle .= '</xf>';
			} else {
				$sStyle .= '/>';
			}

			$aStyles[] = $sStyle;
		}

		$sStyle = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n"
			. '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">'
			. '<fonts count="'.count($aFonts).'" x14ac:knownFonts="'.count($aFonts).'">'.implode('', $aFonts) .'</fonts>'
			. '<fills count="'.count($aFills).'">'.implode('', $aFills).'</fills>'
			. '<borders count="'.count($aBorders).'">'.implode('', array_keys($aBorders)).'</borders>'
			. '<cellStyleXfs count="1">'
			. '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>'
			.'</cellStyleXfs>'
			. '<cellXfs count="'.count($aStyles).'">'.implode('', $aStyles).'</cellXfs>'
			. '<cellStyles count="1">'
			. '<cellStyle name="Standard" xfId="0" builtinId="0"/>'
			. '</cellStyles>'
			. '<dxfs count="0"/>'
			. '<tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleMedium4"/>'
			. '</styleSheet>'
		;

		return $sStyle;
	}

	/**
	 * Generate shared strings xml document.
	 *
	 * @return string
	 */
	protected function sharedStringsXml() {
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
	 * generate worksheet relations, especially for hyperlinks at the moment.
	 *
	 * @param Sheet $oSheet
	 *
	 * @return string
	 */
	protected function worksheetRelationsXml(Sheet &$oSheet) {

		$sXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
			  . '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
		;
		foreach($this->aHyperlinks[$oSheet->getIdentifier()] as $sLink => $iNum) {
			$sXml .= '<Relationship Id="rId' . $iNum . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="' . $sLink . '" TargetMode="External"/>';
		}


		return $sXml . '</Relationships>';
	}

#pragma mark - abstract methods to save and create documents

	/**
	 * This method will save a document to a file and returns if file
	 * could saved or not.
	 *
	 * @param string $sFilename
	 * @throws \LogicException
	 * @return bool
	 */
	public function saveToFile($sFilename) {


		$this->checkNumberOfSheetsInWorkbook();


		// create zip file
		$oZip  = new \ZipArchive;
		$oRes = $oZip->open($sFilename, \ZipArchive::CREATE);

		if (true !== $oRes) {
			throw new \LogicException('Excellence could not create Excel workbook.');
		}

		$sSheetXml = $this->createSheetXml();

		foreach($this->aSheets as $iSheet => $oSheet) {

			/** @var Sheet $oSheet */
			$this->createSheetDataXml($oSheet, $iSheet);

			if ($this->hasHyperlinks($oSheet)) {
				$oZip->addFromString('xl' . DIRECTORY_SEPARATOR . 'worksheets' . DIRECTORY_SEPARATOR . '_rels' . DIRECTORY_SEPARATOR . $oSheet->getIdentifier() . '.xml.rels', $this->worksheetRelationsXml($oSheet));
			}
		}

		// add relations
		$oZip->addFromString('_rels' . DIRECTORY_SEPARATOR . '.rels', $this->documentRelationsXml());

		// add content types
		$oZip->addFromString('[Content_Types].xml', $this->documentContentTypesXml());

		// add app.xml
		$oZip->addFromString('docProps' . DIRECTORY_SEPARATOR . 'app.xml', $this->workbookAppXml());

		// add core.xml
		$oZip->addFromString('docProps' . DIRECTORY_SEPARATOR . 'core.xml', $this->workbookCoreXml());

		// add workbook relations
		$oZip->addFromString('xl' . DIRECTORY_SEPARATOR . '_rels' . DIRECTORY_SEPARATOR . 'workbook.xml.rels', $this->workbookRelationsXml());


		// add shared strings
		if (!empty($this->aSharedStrings)) {
			$oZip->addFromString('xl' . DIRECTORY_SEPARATOR . 'sharedStrings.xml', $this->sharedStringsXml());
		}

		// add styles
		$oZip->addFromString('xl' . DIRECTORY_SEPARATOR . 'styles.xml', $this->workbookStylesXml());

		// add workbook information
		$oZip->addFromString('xl' . DIRECTORY_SEPARATOR . 'workbook.xml', $sSheetXml);

		foreach($this->aSheetData as $sSheetId => $sDom) {
			$oZip->addFromString('xl' . DIRECTORY_SEPARATOR . 'worksheets' . DIRECTORY_SEPARATOR . $sSheetId . '.xml', $sDom);
		}


		$oZip->close();

		return true;
	}


} 