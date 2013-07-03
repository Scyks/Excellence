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
class PerformanceDataSource extends DataSource implements WorkbookDelegate, DataDelegate {

	function cord($iColumn, $iRow) {

		for($sReturn = ""; $iColumn >= 0; $iColumn = intval($iColumn / 26) - 1) {
			$sReturn = chr($iColumn%26 + 0x41) . $sReturn;
		}

		return $sReturn . $iRow;

	}

	/**
	 * construction - load data
	 */
	public function __construct() {
		$this->aSheets[] = new Sheet('sheet1', 'Sheet 1');

		for ($iRow = 0; $iRow < 100; $iRow++) {
			for ($iColumn = 0; $iColumn < 20; $iColumn += 4) {

				$this->aData['sheet1'][$iRow][$iColumn] = 'test column value';
				$this->aData['sheet1'][$iRow][$iColumn+1] = 22;
				$this->aData['sheet1'][$iRow][$iColumn+2] = 0.9;
				$this->aData['sheet1'][$iRow][$iColumn+3] = '=SUM(' . $this->cord($iColumn+1, $iRow+1) . ':' . $this->cord($iColumn+2, $iRow+1) . ')';
			}

		}

		//echo $iRow . ' / ' . $iColumn . "\n";


	}

}