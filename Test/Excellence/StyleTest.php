<?php
/**
 * @author        Ronald Marske <scyks@ceow.de>
 * @filesource    Test/Excellence/StyleTest.php
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
use Excellence\Style;

/**
 * @group Style
 */
class StyleTest extends \PHPUnit_Framework_TestCase {

#pragma mark - creations

	/**
	 * @return Style
	 */
	public function makeStyle() {

		return new Style();
	}

#pragma mark - dataProvider

	public function dataProviderSetter() {
		return array(
			array('setFont', 'sFont', 'foobar'),
			array('setFontSize', 'iFontSize', 11.0),
			array('setBold', 'bBold', false),
			array('setBold', 'bBold', true),
			array('setItalic', 'bItalic', false),
			array('setItalic', 'bItalic', true),
			array('setUnderline', 'bUnderline', false),
			array('setUnderline', 'bUnderline', true),
			array('setColor', 'sColor', 'FFCC00'),
			array('setBackgroundColor', 'sBackground', 'FFCC00'),
			array('setHorizontalAlignment', 'sHorizontalAlignment', Style::ALIGN_CENTER),
			array('setVerticalAlignment', 'sVerticalAlignment', Style::ALIGN_TOP),
			array('setWidth', 'fWidth', 20),
			array('setHeight', 'fHeight', 20),
		);
	}

	public function dataProviderFluent() {
		return array(
//			      method     param
			array('setFont', 'foobar'),
			array('setFontSize', 11.0),
			array('setBold', null),
			array('setItalic', null),
			array('setUnderline', null),
			array('setColor', 'FFCC00'),
			array('setBackgroundColor', 'FFCC00'),
			array('setHorizontalAlignment', Style::ALIGN_CENTER),
			array('setVerticalAlignment', Style::ALIGN_TOP),
			array('setBorder', array(Style::BORDER_DOTTED, Style::BORDER_ALIGN_ALL, 'ffffff')),
			array('setWidth', 20),
			array('setHeight', 20),
		);
	}

	public function dataProviderHas() {
		return array(
//			       has       setter     value     default
			array('hasFont', 'setFont', 'foobar', false),
			array('hasFontSize', 'setFontSize', 11.0, false),
			array('isBold', 'setBold', true, false),
			array('isItalic', 'setItalic', true, false),
			array('hasUnderline', 'setUnderline', true, false),
			array('hasColor', 'setColor', 'FFCC00', false),
			array('hasBackgroundColor', 'setBackgroundColor', 'FFCC00', false),
			array('hasHorizontalAlignment', 'setHorizontalAlignment', Style::ALIGN_CENTER, false),
			array('hasVerticalAlignment', 'setVerticalAlignment', Style::ALIGN_TOP, null),
			array('hasBorder', 'setBorder', array(Style::BORDER_DOTTED, Style::BORDER_ALIGN_ALL, 'ffffff'), false),
			array('hasWidth', 'setWidth', 20, null),
			array('hasHeight', 'setHeight', 20, null),
		);
	}

	public function dataProviderGetter() {
		return array(
//			       getter    setter     value     default
			array('getFont', 'setFont', 'foobar', null),
			array('getFontSize', 'setFontSize', 11.0, null),
			array('getBackgroundColor', 'setBackgroundColor', 'FFCC00', null),
			array('getHorizontalAlignment', 'setHorizontalAlignment', Style::ALIGN_CENTER, null),
			array('getVerticalAlignment', 'setVerticalAlignment', Style::ALIGN_TOP, null),
			array('getWidth', 'setWidth', 20, null),
			array('getHeight', 'setHeight', 20, null),
		);
	}

#pragma mark - identifier

	/**
	 * @test
	 */
	public function getId_createIdentifier_returnMd5String() {
		$this->assertSame('34e7a020e74217423d5f790aac0280de', $this->makeStyle()->getId());
	}

	/**
	 * @test
	 */
	public function getId_idCreated_returnCreeatedMd5String() {
		$oStyle = $this->makeStyle();

		$oReflection = new \ReflectionClass($oStyle);
		$oId = $oReflection->getProperty('sId');
		$oId->setAccessible(true);
		$oId->setValue($oStyle, 'foobar');

		$this->assertSame('foobar', $oStyle->getId());
	}

#pragma mark - setter/getter/has tests

	/**
	 * @param $sSetterMethod
	 * @param $sStoreVariableName
	 * @param $value
	 *
	 * @test
	 * @dataProvider dataProviderSetter
	 */
	public function setterTest_paramSetted_valueStored($sSetterMethod, $sStoreVariableName, $value) {
		$oStyle = $this->makeStyle();

		call_user_func_array(array($oStyle, $sSetterMethod), (array) $value);

		$this->assertAttributeEquals($value, $sStoreVariableName, $oStyle);
	}

	/**
	 * @param $sHasMethod
	 * @param $sSetterMethod
	 * @param $value
	 * @param $bDefault
	 *
	 * @test
	 * @dataProvider dataProviderHas
	 */
	public function hasTest_NoParamSetted_returnsDefault($sHasMethod, $sSetterMethod, $value, $bDefault) {
		$oStyle = $this->makeStyle();

		$this->assertEquals($bDefault, $oStyle->{$sHasMethod}());
	}

	/**
	 * @param $sHasMethod
	 * @param $sSetterMethod
	 * @param $value
	 * @param $bDefault
	 *
	 * @test
	 * @dataProvider dataProviderHas
	 */
	public function hasTest_paramSetted_returnsDefault($sHasMethod, $sSetterMethod, $value, $bDefault) {
		$oStyle = $this->makeStyle();

		call_user_func_array(array($oStyle, $sSetterMethod), (array) $value);

		$this->assertEquals(true, $oStyle->{$sHasMethod}());
	}

	/**
	 * @param $sGetterMethod
	 * @param $sSetterMethod
	 * @param $value
	 * @param $bDefault
	 *
	 * @test
	 * @dataProvider dataProviderGetter
	 */
	public function getterTest_NoParamSetted_returnsDefault($sGetterMethod, $sSetterMethod, $value, $bDefault) {
		$oStyle = $this->makeStyle();

		$this->assertEquals($bDefault, $oStyle->{$sGetterMethod}());
	}

	/**
	 * @param $sGetterMethod
	 * @param $sSetterMethod
	 * @param $value
	 * @param $bDefault
	 *
	 * @test
	 * @dataProvider dataProviderGetter
	 */
	public function getterTest_paramSetted_returnsDefault($sGetterMethod, $sSetterMethod, $value, $bDefault) {
		$oStyle = $this->makeStyle();

		call_user_func_array(array($oStyle, $sSetterMethod), (array) $value);

		$this->assertEquals($value, $oStyle->{$sGetterMethod}());
	}

	/**
	 * @param $sMethod
	 * @param $params
	 *
	 * @test
	 * @dataProvider dataProviderFluent
	 */
	public function fluentInterfaceTest_callMethod_returnsObject($sMethod, $params = array()) {
		$oStyle = $this->makeStyle();

		$bCompare = call_user_func_array(array($oStyle, $sMethod), (array) $params);


		$this->assertSame($oStyle, $bCompare);
	}

#pragma mark - color checks

	/**
	 * @test
	 * @expectedException \InvalidArgumentException
	 * @expectedExceptionMessage Please provide a hexadecimal color code.
	 */
	public function checkColor_wrongFormat_throwsException() {
		$oStyle = $this->makeStyle();

		$oStyle->setColor('foobar');
	}

	/**
	 * @test
	 */
	public function setColor_lowerCaseColor_storeUppercaseColor() {
		$oStyle = $this->makeStyle();

		$oStyle->setColor('ffffff');

		$this->assertSame('FFFFFF', $oStyle->getColor());
	}

	/**
	 * @test
	 * @expectedException \InvalidArgumentException
	 * @expectedExceptionMessage Please provide a hexadecimal color code.
	 */
	public function checkBackgroundColor_wrongFormat_throwsException() {
		$oStyle = $this->makeStyle();

		$oStyle->setBackgroundColor('foobar');
	}

	/**
	 * @test
	 */
	public function setBackgroundColor_lowerCaseColor_storeUppercaseColor() {
		$oStyle = $this->makeStyle();

		$oStyle->setBackgroundColor('ffffff');

		$this->assertSame('FFFFFF', $oStyle->getBackgroundColor());
	}

#pragma mark - alignment

	/**
	 * @test
	 * @expectedException \InvalidArgumentException
	 * @expectedExceptionMessage Horizontal alignment allows only "center, general, justify, left, right".
	 */
	public function setHorizontalSAlignment_wrongAlignment_throwsException() {
		$oStyle = $this->makeStyle();

		$oStyle->setHorizontalAlignment('foobar');
	}

	/**
	 * @test
	 * @expectedException \InvalidArgumentException
	 * @expectedExceptionMessage Vertical alignment allows only "center, top, justify, bottom".
	 */
	public function setVerticalAlignment_wrongAlignment_throwsException() {
		$oStyle = $this->makeStyle();

		$oStyle->setVerticalAlignment('foobar');
	}

#pragma mark - borders

	/**
	 * @test
	 * @expectedException \InvalidArgumentException
	 * @expectedExceptionMessage Border style "foobar" is not allowed. Please provide a valid style.
	 */
	public function setBorder_wrongBorder_throwsException() {
		$oStyle = $this->makeStyle();

		$oStyle->setBorder('foobar', '', '');
	}

	/**
	 * @test
	 * @expectedException \InvalidArgumentException
	 * @expectedExceptionMessage Please provide a valid border alignment style.
	 */
	public function setBorder_wrongAlignment_throwsException() {
		$oStyle = $this->makeStyle();

		$oStyle->setBorder(Style::BORDER_DOTTED, 'foobar', '');
	}

	/**
	 * @test
	 * @expectedException \InvalidArgumentException
	 * @expectedExceptionMessage Please provide a hexadecimal color code.
	 */
	public function setBorder_wrongColor_throwsException() {
		$oStyle = $this->makeStyle();

		$oStyle->setBorder(Style::BORDER_DOTTED, Style::BORDER_ALIGN_ALL, 'foobar');
	}

	/**
	 * @test
	 */
	public function setBorder_correctParams_borderStored() {
		$oStyle = $this->makeStyle();
		$oStyle->setBorder(Style::BORDER_DOTTED, Style::BORDER_ALIGN_ALL, 'ffffff');
		
		$aCompare = array (
			1 => array (
				'style' => 'dotted',
		        'color' => 'FFFFFF',
		    ),
		    2 => array (
				'style' => 'dotted',
		        'color' => 'FFFFFF',
		    ),
		    4 => array (
				'style' => 'dotted',
		        'color' => 'FFFFFF',
		    ),
		    8 => array (
				'style' => 'dotted',
		        'color' => 'FFFFFF',
		    ),
		);
		
		$this->assertSame($aCompare, $oStyle->getBorder());
	}
}
