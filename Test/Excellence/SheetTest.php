<?php
/**
 * @author        Ronald Marske <scyks@ceow.de>
 * @filesource    Test/Excellence/SheetTest.php
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
use Test\TestCase;
/**
 * Class SheetTest
 * @group Sheet
 * @package Test\Excellence
 */
class SheetTest extends TestCase {

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

#pragma mark - dataProvider

	public function dataProviderIllegalIds() {
		return array(
			array('foo bar'),
			array('foo/&%$§'),
			array('foo:'),
			array(':bar:'),
			array(':bar:'),
			array(':bar<'),
			array(''),
		);
	}

#pragma mark - construction

	/**
	 * @test
	 * @dataProvider dataProviderIllegalIds
	 * @expectedException \InvalidArgumentException
	 * @expectedExceptionMessage Sheet identifier does only contain following signs (a-z, 0-9, _, -).
	 */
	public function __construct_IllegalIdentifiers_throwsException($sId) {
		$this->makeSheet($sId);
	}


	/**
	 * @test
	 */
	public function __construct_provideStringIdentifier_identifierStored() {

		$oSheet = $this->makeSheet('foo');
		$this->assertAttributeSame('foo', 'sIdentifier', $oSheet);
	}

	/**
	 * @test
	 */
	public function __construct_nullAsName_nameEqualsNull() {

		$oSheet = $this->makeSheet('foo');
		$this->assertAttributeSame(null, 'sName', $oSheet);
	}

	/**
	 * @test
	 */
	public function __construct_emptyStringAsName_nameEqualsNull() {

		$oSheet = $this->makeSheet('foo', '');
		$this->assertAttributeSame(null, 'sName', $oSheet);
	}

	/**
	 * @test
	 */
	public function __construct_stringAsName_nameEqualsFoobar() {

		$oSheet = $this->makeSheet('foo', 'Foobar');
		$this->assertAttributeSame('Foobar', 'sName', $oSheet);
	}


#pragma mark - getIdentifier

	/**
	 * @test
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
	 */
	public function getIdentifier_constructionIdentifier_returnsSheet() {
		$oSheet = $this->makeSheet();
		$this->assertSame('sheet', $oSheet->getIdentifier());
	}

#pragma mark - getName

	/**
	 * @test
	 */
	public function getName_noNameProvided_returnsNull() {

		$oSheet = $this->makeSheet();
		$this->assertNull($oSheet->getName());
	}

	/**
	 * @test
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
	 */
	public function getName_constructionIdentifier_returnsSheet() {
		$oSheet = $this->makeSheet('scheet', 'Foobar');
		$this->assertSame('Foobar', $oSheet->getName());
	}

#pragma mark - hasName


	/**
	 * @test
	 */
	public function hasName_noNameProvided_returnsFalse() {

		$oSheet = $this->makeSheet();
		$this->assertFalse($oSheet->hasName());
	}

	/**
	 * @test
	 */
	public function hasName_modifiedIdentifierVariable_returnsFoobar() {

		$oSheet = $this->makeSheet();

		$oReflectionClass = new \ReflectionClass($oSheet);

		$oProperty = $oReflectionClass->getProperty('sName');
		$oProperty->setAccessible(true);
		$oProperty->setValue($oSheet, 'foobar');

		$this->assertTrue($oSheet->hasName());
	}

	/**
	 * @test
	 */
	public function hasName_constructionIdentifier_returnsSheet() {
		$oSheet = $this->makeSheet('scheet', 'Foobar');
		$this->assertTrue($oSheet->hasName());
	}

#pragma mark - markFirstRowAsFixed

	/**
	 * @test
	 */
	public function firstRowAsFixed_defaultValue_storedAsFalse() {
		$oSheet = $this->makeSheet();

		$this->assertAttributeEquals(false, 'bFirstRowFixed', $oSheet);
	}

	/**
	 * @test
	 */
	public function setFirstRowAsFixed_setToFalse_storedAsFalse() {
		$oSheet = $this->makeSheet();

		$oSheet->setFirstRowAsFixed(false);

		$this->assertAttributeEquals(false, 'bFirstRowFixed', $oSheet);
	}

	/**
	 * @test
	 * @dataProvider dataProviderInvalidBoolean
	 * @expectedException \InvalidArgumentException
	 * @expectedExceptionMessage Please provide a boolean value to "Excellence\Sheet::setFirstRowAsFixed".
	 */
	public function setFirstRowAsFixed_noBoolean_throwException($value) {
		$oSheet = $this->makeSheet();

		$oSheet->setFirstRowAsFixed($value);
	}

	/**
	 * @test
	 */
	public function setFirstRowAsFixed_setToTrue_storedAsTrue() {
		$oSheet = $this->makeSheet();

		$oSheet->setFirstRowAsFixed(true);

		$this->assertAttributeEquals(true, 'bFirstRowFixed', $oSheet);
	}

	/**
	 * @test
	 */
	public function setFirstRowAsFixed_checkFluentInterface_returnSelf() {
		$oSheet = $this->makeSheet();

		$this->assertSame($oSheet, $oSheet->setFirstRowAsFixed(true));
	}

#pragma mark - isFirstRowFixed

	/**
	 * @test
	 */
	public function isFirstRowFixed_defaultValue_returnsFalse() {
		$oSheet = $this->makeSheet();

		$this->assertFalse($oSheet->isFirstRowFixed());
	}

	/**
	 * @test
	 */
	public function isFirstRowFixed_settedToFalse_returnsFalse() {
		$oSheet = $this->makeSheet();
		$oSheet->setFirstRowAsFixed(false);

		$this->assertFalse($oSheet->isFirstRowFixed());
	}

	/**
	 * @test
	 */
	public function isFirstRowFixed_settedToTrue_returnsTrue() {
		$oSheet = $this->makeSheet();
		$oSheet->setFirstRowAsFixed(true);

		$this->assertTrue($oSheet->isFirstRowFixed());
	}

	/**
	 * @test
	 */
	public function isFirstRowFixed_checkIfReturnValueIsTypeCastedToBool_returnsTrue() {
		$oSheet = $this->makeSheet();

		$oReflectionClass = new \ReflectionClass($oSheet);

		$oProperty = $oReflectionClass->getProperty('bFirstRowFixed');
		$oProperty->setAccessible(true);
		$oProperty->setValue($oSheet, 'foobar');

		$this->assertTrue($oSheet->isFirstRowFixed());
	}

}