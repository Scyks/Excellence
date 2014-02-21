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
use Excellence\Delegates\MergeableDelegate;
use Excellence\Delegates\StylableDelegate;
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
	 * @var null|Style
	 */
	private $oStandardStyle = null;

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

		$this->createStandardStyles();

	}

#pragma mark - styles

	/**
	 * this method will create standard style information
	 */
	private function createStandardStyles() {
		$this->oStandardStyle = new Style();
		$this->oStandardStyle
			->setFont('Calibri')
			->setFontSize(12)
			->setColor('000000')
		;
	}

	/**
	 * returns standard style information
	 * @return Style
	 */
	public function getStandardStyles() {
		return $this->oStandardStyle;
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

#pragma mark - coordinates

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
	public function getCoordinatesByColumnAndRow($iColumn, $iRow = null) {

		for($sReturn = ""; $iColumn >= 0; $iColumn = intval($iColumn / 26) - 1) {
			$sReturn = chr($iColumn%26 + 0x41) . $sReturn;
		}

		// if there is no row, only the letter will returned
		if (null === $iRow) {
			return $sReturn;
		}

		return $sReturn . $iRow;

	}

}