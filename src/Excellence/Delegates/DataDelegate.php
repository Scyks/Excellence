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

namespace Excellence\Delegates;

use Excellence\Sheet;
use Excellence\Workbook;

/**
 * Data source delegate interface to provide functionality to create data
 * for a sheet of an Excel document
 * .
 * @package Excellence\Delegates
 */
interface DataDelegate {

	/**
	 * returns an integer value how many rows a sheet for workbook contains.
	 *
	 * @param Workbook $oWorkbook
	 * @param Sheet $oSheet
	 *
	 * @return int
	 */
	public function numberOfRowsInSheet(Workbook $oWorkbook, Sheet $oSheet);

	/**
	 * returns an integer value how many columns a sheet for workbook contains.
	 *
	 * @param Workbook $oWorkbook
	 * @param Sheet $oSheet
	 *
	 * @return int
	 */
	public function numberOfColumnsInSheet(Workbook $oWorkbook, Sheet $oSheet);

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
	public function valueForRowAndColumn(Workbook $oWorkbook, Sheet $oSheet, $iRow, $iColumn);
}