<?php
/**
 * @author		  Ronald Marske <r.marske@secu-ring.de>
 * @filesource	  src/Excellence/Writer/AbstractWriter.php
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

namespace Excellence\Writer;

use \Excellence\Workbook;
use Excellence\Delegates\WorkbookDelegate;

/**
 * Abstraction of document writer class
 * @package Excellence\Writer
 */
abstract class AbstractWriter {

	/**
	 * @var Workbook
	 */
	private $oWorkbook = null;

	/**
	 * Construction of excel document writer for Excellence
	 * @param Workbook $oWorkbook
	 */
	public function __construct(Workbook $oWorkbook) {
		$this->oWorkbook = $oWorkbook;
	}

	/**
	 * @return Workbook
	 */
	public function getWorkbook() {
		return $this->oWorkbook;
	}

#pragma mark - get delegate

	/**
	 * @return WorkbookDelegate
	 */
	public function getDelegate() {
		return $this->getWorkbook()->getDelegate();
	}

#pragma mark - abstract methods

	/**
	 * This method will save a document to a file and returns if file
	 * could saved or not.
	 *
	 * @param string $sFilename
	 * @return bool
	 */
	abstract public function saveToFile($sFilename);

} 