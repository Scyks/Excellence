<?php
/**
 * @author        Ronald Marske <scyks@ceow.de>
 * @filesource    src/Execllence/Delegate/StylableDelegate.php
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
use Excellence\Style;
use Excellence\Workbook;

interface StylableDelegate {

	/**
	 * returns a Style definition that will be used as default for all cells.
	 *
	 * @param Workbook $oWorkbook
	 * @param Sheet    $oSheet
	 *
	 * @return Style
	 */
	public function getStandardStyle(Workbook $oWorkbook, Sheet $oSheet);

	/**
	 * returns Style definition for a specific cell by given identifier
	 *
	 * @param Workbook $oWorkbook
	 * @param Sheet    $oSheet
	 * @param int      $iColumn
	 * @param int      $iRow
	 *
	 * @return Style
	 */
	public function getStyleForColumnAndRow(Workbook $oWorkbook, Sheet $oSheet, $iColumn, $iRow);
}