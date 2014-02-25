<?php
/**
 * @author        Ronald Marske <scyks@ceow.de>
 * @filesource    src/Excellence/Delegates/MergeableDelegate.php
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

use Excellence\Workbook;
use Excellence\Sheet;

interface LinkableDelegate {

	/**
	 * this method will return true when a specific column in a specific
	 * row has an hyperlink. If this column has any, the method getLinkForColumnAndRow
	 * will called.
	 *
	 * @param Workbook $oWorkbook
	 * @param Sheet    $oSheet
	 * @param int      $iRow
	 * @param int      $iColumn
	 *
	 * @return boolean
	 */
	public function hasLinkForColumnAndRow(Workbook $oWorkbook, Sheet $oSheet, $iRow, $iColumn);

	/**
	 * this method will return a hyperlink (url) for a specific column in a specific
	 * row. This method is only been called, when getLinkForColumnAndRow returns
	 * true.
	 *
	 * @param Workbook $oWorkbook
	 * @param Sheet    $oSheet
	 * @param int      $iRow
	 * @param int      $iColumn
	 *
	 * @return string
	 */
	public function getLinkForColumnAndRow(Workbook $oWorkbook, Sheet $oSheet, $iRow, $iColumn);
}