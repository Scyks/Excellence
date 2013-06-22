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
	 * @var \DomDocument|null
	 */
	private $oSheets;

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
		}

		// unset variables
		unset($oSheets, $oXmlSheet, $iSheet, $iSheets);

	}
}