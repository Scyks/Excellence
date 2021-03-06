<?php
/**
 * @author        Ronald Marske <scyks@ceow.de>
 * @filesource    src/Excellence/Sheet.php
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

/**
 * Workbook Sheet Document. This class represent a sheet document
 * like Excel Sheet.
 *
 * @package Excellence
 */
class Sheet {

#pragma mark - member variables

	/**
	 * sheet identifier
	 * @var string
	 */
	private $sIdentifier;

	/**
	 * Name of this sheet
	 * @var string
	 */
	private $sName;

	/**
	 * flag to mark first row as fixed
	 * @var bool
	 */
	private $bFirstRowFixed = false;

	/**
	 * position of row, at which row it will be fixed
	 * @var int
	 */
	private $iFixRowPosition = 1;

#pragma mark - construction

	/**
	 * Creates a new Sheet document with given identifier. Identifier have to
	 * be a non empty string. there can also provided a name for this sheet
	 * document. this name will be displayed in Excel workbook at the bottom.
	 *
	 * @param string $sIdentifier
	 * @param null $sName
	 * @throws \InvalidArgumentException
	 */
	public function __construct($sIdentifier, $sName = null) {

		// make sure identifier isn't empty
		if(!preg_match('/^[a-z0-9_-]+$/i', $sIdentifier)) {
			throw new \InvalidArgumentException('Sheet identifier does only contain following signs (a-z, 0-9, _, -).');
		}

		// set name to null if empty
		if(null !== $sName && empty($sName)) {
			$sName = null;
		}

		$this->sIdentifier = $sIdentifier;
		$this->sName = $sName;
	}

#pragma mark - identifier

	/**
	 * return sheet identifier to identify a workbook. This
	 * should be unique in an application.
	 * @return string
	 */
	public function getIdentifier() {
		return $this->sIdentifier;
	}

#pragma mark - name

	/**
	 * return sheet name
	 *
	 * @return string
	 */
	public function getName() {
		return $this->sName;
	}

	/**
	 * check if a sheet has a name defined
	 * @return bool
	 */
	public function hasName() {
		return (0 < strlen($this->sName));
	}

#pragma mark - mark first row as fixed

	/**
	 * this method will mark the first line as fixed, that i can't be
	 * scrolled.
	 *
	 * @param bool $bFixed
	 * @return Sheet
	 * @throws \InvalidArgumentException
	 */
	public function setFirstRowAsFixed($bFixed, $iRow = 1) {

		if (!is_bool($bFixed)) {
			throw new \InvalidArgumentException(sprintf('Please provide a boolean value to "%s".', __METHOD__));
		}

		$this->bFirstRowFixed = $bFixed;
		$this->iFixRowPosition = $iRow;

		return $this;
	}

	public function getRowFixPosition() {
		return $this->iFixRowPosition;
	}

	/**
	 * Check method if first row is marked as fixed
	 * @return bool
	 */
	public function isFirstRowFixed() {
		return (bool) $this->bFirstRowFixed;
	}
}