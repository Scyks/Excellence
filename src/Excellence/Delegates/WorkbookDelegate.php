<?php
/**
 * @author        Ronald Marske <scyks@ceow.de>
 * @filesource    src/Excellence/Delegates/WorkbookDelegate.php
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

use \Excellence\Workbook;
use \Excellence\Sheet;

/**
 * Workbook delegate interface to provide minimum functionality to
 * create an Excel document.
 *
 * @package Excellence\Delegates
 */
interface WorkbookDelegate {

	/**
	 * returns integer value of how many sheets this workbook
	 * will contain.
	 *
	 * @param Workbook $oWorkbook
	 *
	 * @return int
	 */
	public function numberOfSheetsInWorkbook(Workbook $oWorkbook);

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
	public function getSheetForWorkBook(Workbook $oWorkbook, $iSheetIndex);

	/**
	 * Return a data source instance for given workbook and sheet.
	 * This method will be called n times. N is the number retrieved
	 * by numberOfSheetsInWorkbook.
	 *
	 * @param Workbook $oWorkbook
	 * @param Sheet $oSheet
	 * @return mixed
	 */
	public function dataSourceForWorkbookAndSheet(Workbook $oWorkbook, Sheet $oSheet);
}