<?php
/**
 * @author        Ronald Marske <scyks@ceow.de>
 * @filesource    ExcellenceTest.php
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

namespace Test\Excellence;
use Excellence\Sheet;

class SheetTest extends \PHPUnit_Framework_TestCase {

#pragma mark - creations

	/**
	 * @param string $sIdentifier
	 * @param string $sName
	 *
	 * @return Workbook
	 */
	public function makeSheet($sIdentifier = 'sheet', $sName = null) {

		return new Sheet($sIdentifier, $sName);
	}


#pragma mark - construction

	/**
	 * @test
	 * @group Sheet
	 * @expectedException \InvalidArgumentException
	 * @expectedExceptionMessage Sheet identifier have to be a non empty string value.
	 */
	public function __construct_provideEmptyIdentifier_throwsException() {
		$this->makeSheet('');
	}


	/**
	 * @test
	 * @group Sheet
	 */
	public function __construct_provideStringIdentifier_identifierStored() {

		$oSheet = $this->makeSheet('foo');
		$this->assertAttributeSame('foo', 'sIdentifier', $oSheet);
	}

	/**
	 * @test
	 * @group Sheet
	 */
	public function __construct_nullAsName_nameEqualsNull() {

		$oSheet = $this->makeSheet('foo');
		$this->assertAttributeSame(null, 'sName', $oSheet);
	}

	/**
	 * @test
	 * @group Sheet
	 */
	public function __construct_emptyStringAsName_nameEqualsNull() {

		$oSheet = $this->makeSheet('foo', '');
		$this->assertAttributeSame(null, 'sName', $oSheet);
	}

	/**
	 * @test
	 * @group Sheet
	 */
	public function __construct_stringAsName_nameEqualsFoobar() {

		$oSheet = $this->makeSheet('foo', 'Foobar');
		$this->assertAttributeSame('Foobar', 'sName', $oSheet);
	}


#pragma mark - getIdentifier

	/**
	 * @test
	 * @group Sheet
	 */
	public function getIdentifier_modifiedIdentifierVariable_returnsFoobar() {

		$oSheet = $this->makeSheet();

		$oReflectionClass = new \ReflectionClass($oSheet);

		$oProperty = $oReflectionClass->getProperty('sIdentifier');
		$oProperty->setAccessible(true);
		$oProperty->setValue($oSheet, 'foobar');

		$this->assertSame('foobar', $oSheet->getIdentifier());
	}

	/**
	 * @test
	 * @group Sheet
	 */
	public function getIdentifier_constructionIdentifier_returnsSheet() {
		$oSheet = $this->makeSheet();
		$this->assertSame('sheet', $oSheet->getIdentifier());
	}

#pragma mark - getName

	/**
	 * @test
	 * @group Sheet
	 */
	public function getName_noNameProvided_returnsNull() {

		$oSheet = $this->makeSheet();
		$this->assertNull($oSheet->getName());
	}

	/**
	 * @test
	 * @group Sheet
	 */
	public function getName_modifiedIdentifierVariable_returnsFoobar() {

		$oSheet = $this->makeSheet();

		$oReflectionClass = new \ReflectionClass($oSheet);

		$oProperty = $oReflectionClass->getProperty('sName');
		$oProperty->setAccessible(true);
		$oProperty->setValue($oSheet, 'foobar');

		$this->assertSame('foobar', $oSheet->getName());
	}

	/**
	 * @test
	 * @group Sheet
	 */
	public function getName_constructionIdentifier_returnsSheet() {
		$oSheet = $this->makeSheet('scheet', 'Foobar');
		$this->assertSame('Foobar', $oSheet->getName());
	}

}