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

use Excellence\Delegates\MergeableDelegate;
use Excellence\Delegates\WorkbookDelegate;
use \Excellence\Workbook;
use \Excellence\Sheet;

/**
 * Stub Data Source for Tests
 * @package Test\Excellence\Stub
 */
class MergeableDataSource extends DataSource implements MergeableDelegate, WorkbookDelegate {

	/**
	 * construction - load data
	 */
	public function __construct() {
		$this->aSheets[] = new Sheet('sheet1', 'Sheet 1');

		$this->aData['sheet1'][0][] = 'row1col1';	// A0
		$this->aData['sheet1'][0][] = 'B1:C1';		// B0
		$this->aData['sheet1'][1][] = 'row2col1';	// A1
		$this->aData['sheet1'][1][] = 'B2:C3';		// B1
		$this->aData['sheet1'][2][] = 'row3col1';	// A2
		$this->aData['sheet1'][3][] = 'A4:A5';		// A3

	}

#pragma mark - MergeableDelegate implementation

	/**
	 * Returns a string merge definition as string. A merge string will
	 * contain two cell coordinates separated by a colon. Workbook's API
	 * provide a method to retrieve a cell coordinate by column and row.
	 *
	 * - Workbook::getCoordinatesByColumnAndRow(int $iColumn, int $iRow)
	 *
	 * Format examples:
	 * - A1:B1
	 * - A1:B2
	 *
	 * @param Workbook $oWorkbook
	 * @param int $iColumn
	 * @param int $iRow
	 *
	 * @return string
	 */
	public function mergeByColumnAndRow(Workbook $oWorkbook, $iColumn, $iRow) {
		if (1 == $iColumn && 0 == $iRow) {
			return 'B1:C1';
		} elseif (1 == $iColumn && 1 == $iRow) {
			return 'B2:C3';
		} elseif (0 == $iColumn && 3 == $iRow) {
			return 'A3:A4';
		}

		return null;
	}

}