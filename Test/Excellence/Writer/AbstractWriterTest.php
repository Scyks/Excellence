<?php
/**
 * @author		  Ronald Marske <r.marske@secu-ring.de>
 * @filesource	  Test/Excellence/Writer/AbstractWriterTesst.php
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

namespace Test\Excellence\Writer;

use Excellence\Workbook;
use Excellence\Writer\AbstractWriter;
use Test\TestCase;

/**
 * AbstractWriter Test class
 * @package Test\Excellence\Writer
 * @group AbstractWriter
 */
class AbstractWriterTest extends TestCase {

#pragma mark - creation

	/**
	 * @param Workbook $oWorkbook
	 * @return AbstractWriter
	 */
	private function createWriter(Workbook $oWorkbook = null) {

		if (!$oWorkbook instanceof Workbook)
			$oWorkbook = $this->makeWorkbook();

		return $this->getMockForAbstractClass('\Excellence\Writer\AbstractWriter', array($oWorkbook));
	}

#pragma mark - construction

	/**
	 * @test
	 */
	public function __construct_givenWorkbook_workbookStored() {
		$oWorkbook = $this->makeWorkbook();

		$oWriter = $this->createWriter($oWorkbook);

		$this->assertAttributeEquals($oWorkbook, 'oWorkbook', $oWriter);
	}

#pragma mark - getWorkbook

	/**
	 * @test
	 */
	public function getWorkbook_givenWorkbook_workbookReturned() {
		$oWorkbook = $this->makeWorkbook();

		$oWriter = $this->createWriter($oWorkbook);

		$this->assertEquals($oWorkbook, $oWriter->getWorkbook());
	}
#pragma mark - getDelegate

	/**
	 * @test
	 */
	public function getDelegate_getDelegateFromWorkbook_DelegateReturned() {
		$oWorkbook = $this->makeWorkbook();

		$oWriter = $this->createWriter($oWorkbook);

		$this->assertEquals($oWorkbook->getDelegate(), $oWriter->getDelegate());
	}

}