<?php
/**
 * @author        Ronald Marske <scyks@ceow.de>
 * @filesource    DataSource.php
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

namespace Test\Excellence\Stub;

use \Excellence\Delegates\WorkbookDelegate;
use \Excellence\Delegates\DataDelegate;
use \Excellence\Workbook;
use \Excellence\Sheet;

/**
 * Stub Data Source for Tests
 * @package Test\Excellence\Stub
 */
class DataSource implements WorkbookDelegate, DataDelegate {

	/**
	 * sheets
	 * @var array
	 */
	protected $aSheets = array();

	/**
	 * sheet data
	 * @var array
	 */
	protected $aData = array();

	/**
	 * construction - load data
	 */
	public function __construct() {
		$this->aSheets[] = new Sheet('sheet1', 'Sheet 1');
		$this->aSheets[] = new Sheet('sheet2', 'Sheet 2');

		$this->aData['sheet1'][0][] = 'row1col1';
		$this->aData['sheet1'][0][] = 42;
		$this->aData['sheet1'][0][] = 42.34;
		$this->aData['sheet1'][1][] = 'row2col1';
		$this->aData['sheet1'][1][] = 42;
		$this->aData['sheet1'][1][] = 42.34;
		$this->aData['sheet1'][2][] = 'row3col1';
		$this->aData['sheet1'][2][] = 42;
		$this->aData['sheet1'][2][] = 42.34;
		$this->aData['sheet1'][3][1] = '=SUM(B1:B3)';
		$this->aData['sheet1'][3][2] = '=SUM(C1:C3)';

		$this->aData['sheet2'] = $this->aData['sheet1'];

	}

#pragma mark - WorkbookDelegate implementation

	/**
	 * returns integer value of how many sheets this workbook
	 * will contain.
	 *
	 * @param Workbook $oWorkbook
	 *
	 * @return int
	 */
	public function numberOfSheetsInWorkbook(Workbook $oWorkbook) {
		return count($this->aSheets);
	}

	/**
	 * Return a Sheet for Workbook by given sheet index. If there
	 * three sheets (numberOfSheetsInWorkbook) available, this
	 * method will called three times by provide sheet index 0, 1
	 * and 2.
	 *
	 * @param Workbook $oWorkbook
	 * @param integer $iSheetIndex
	 *
	 * @return Sheet
	 */
	public function getSheetForWorkBook(Workbook $oWorkbook, $iSheetIndex) {
		return $this->aSheets[$iSheetIndex];
	}

	/**
	 * Return a data source instance for given workbook and sheet.
	 * This method will be called n times. N is the number retrieved
	 * by numberOfSheetsInWorkbook.
	 *
	 * @param Workbook $oWorkbook
	 * @param Sheet $oSheet
	 *
	 * @return mixed
	 */
	public function dataSourceForWorkbookAndSheet(Workbook $oWorkbook, Sheet $oSheet) {
		return $this;
	}

	/**
	 * returns an integer value how many rows a sheet for workbook contains.
	 *
	 * @param Workbook $oWorkbook
	 * @param Sheet $oSheet
	 *
	 * @return int
	 */
	public function numberOfRowsInSheet(Workbook $oWorkbook, Sheet $oSheet) {
		return count($this->aData[$oSheet->getIdentifier()]);
	}

	/**
	 * returns an integer value how many columns a sheet for workbook contains.
	 *
	 * @param Workbook $oWorkbook
	 * @param Sheet $oSheet
	 *
	 * @return int
	 */
	public function numberOfColumnsInSheet(Workbook $oWorkbook, Sheet $oSheet) {
		return count($this->aData[$oSheet->getIdentifier()][0]);
	}

	/**
	 * returns a value for sheet cell by given workbook, sheet, row number and
	 * column number. possible values are integers, floats, doubles and strings.
	 * An Excel function like "=SUM(A1:A4)" will also provided as string.
	 *
	 * @param Workbook $oWorkbook
	 * @param Sheet $oSheet
	 * @param integer $iRow
	 * @param integer $iColumn
	 *
	 * @return string|float|double|int
	 */
	public function valueForRowAndColumn(Workbook $oWorkbook, Sheet $oSheet, $iRow, $iColumn) {
		if (!isset($this->aData[$oSheet->getIdentifier()][$iRow][$iColumn]))
			return null;

		return $this->aData[$oSheet->getIdentifier()][$iRow][$iColumn];
	}

}