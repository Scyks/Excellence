<?php
/**
 * @author        Ronald Marske <scyks@ceow.de>
 * @filesource    src/Excellence/Style.php
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
 * class to style a cell or even a document.
 *
 * @package Excellence
 */
class Style {

#pragma mark - border constants

	const BORDER_NONE					= 'none';
	const BORDER_HAIR					= 'hair';
	const BORDER_THIN					= 'thin';
	const BORDER_MEDIUM					= 'medium';
	const BORDER_THICK					= 'thick';
	const BORDER_DOUBLE					= 'bouble';
	const BORDER_DOTTED					= 'dotted';
	const BORDER_DASH_DOT				= 'dashDot';
	const BORDER_MEDIUM_DASH_DOT		= 'mediumDashDot';
	const BORDER_SLANT_DASH_DOT			= 'slantDashDot';
	const BORDER_DASH_DOT_DOT			= 'dashDotDot';
	const BORDER_MEDIUM_DASH_DOT_DOT	= 'mediumDashDotDot';
	const BORDER_DASHED					= 'dashed';
	const BORDER_MEDIUM_DASHED			= 'mediumDashed';

	const BORDER_ALIGN_TOP				= 1;
	const BORDER_ALIGN_RIGHT			= 2;
	const BORDER_ALIGN_BOTTOM			= 4;
	const BORDER_ALIGN_LEFT				= 8;

#pragma mark - alignment constants

	const ALIGN_CENTER					= 'center';
	const ALIGN_GENERAL					= 'general';
	const ALIGN_JUSTIFY					= 'justify';
	const ALIGN_LEFT					= 'left';
	const ALIGN_RIGHT					= 'right';
	const ALIGN_TOP						= 'top';
	const ALIGN_BOTTOM					= 'bottom';

#pragma mark - member variables

	/**
	 * font family
	 *
	 * @var null|string
	 */
	private $sFont = null;

	/**
	 * font size
	 *
	 * @var null|float
	 */
	private $iFontSize = null;

	/**
	 * bold
	 * @var bool
	 */
	private $bBold = false;

	/**
	 * italic
	 * @var bool
	 */
	private $bItalic = false;

	/**
	 * foreground color
	 *
	 * @var null|string
	 */
	private $sColor = null;

	/**
	 * background color
	 *
	 * @var null|string
	 */
	private $sBackground = null;

	/**
	 * horizontal alignment
	 *
	 * @var null|string
	 */
	private $sHorizontalAlignment = null;

	/**
	 * vertical alignment
	 *
	 * @var null|string
	 */
	private $sVerticalAlignment = null;

	/**
	 * border styles
	 *
	 * @var array
	 */
	private $aBorder = array();

#pragma mark - font handling

	/**
	 * defines a font family
	 * 
	 * @param string $sFont
	 * @return Style
	 */
	public function setFont($sFont) {
		$this->sFont = $sFont;

		return $this;
	}

	/**
	 * return font family
	 * 
	 * @return null|string
	 */
	public function getFont() {
		return $this->sFont;
	}

	/**
	 * defines a font size - will casted to float value
	 * 
	 * @param float $iFontSize
	 * @return Style
	 */
	public function setFontSize($iFontSize) {
		$this->iFontSize = (float) $iFontSize;

		return $this;
	}

	/**
	 * return defined font size
	 * @return float|null
	 */
	public function getFontSize() {
		return $this->iFontSize;
	}

	/**
	 * toggle function for marking font as bold
	 * @param bool $bBold
	 * @return Style
	 */
	public function setBold($bBold = true) {
		$this->bBold = (bool) $bBold;

		return $this;
	}

	/**
	 * check if font is marked as bold
	 * 
	 * @return bool
	 */
	public function isBold() {
		return $this->bBold;
	}
	
	/**
	 * toggle function for marking font as italic
	 * @param bool $bItalic
	 * @return Style
	 */
	public function setItalic($bItalic = true) {
		$this->bItalic = (bool) $bItalic;

		return $this;
	}

	/**
	 * check if font is marked as italic
	 * 
	 * @return bool
	 */
	public function isItalic() {
		return $this->bItalic;
	}
	
#pragma mark - color handling

	/**
	 * defines a hexadecimal color code
	 * 
	 * @param string $sColor
	 * @return Style
	 */
	public function setColor($sColor) {
		$this->checkColor($sColor);
		
		$this->sColor = $sColor;

		return $this;
	}

	/**
	 * return hexadecimal color code
	 * 
	 * @return null|string
	 */
	public function getColor() {
		return $this->sColor;
	}
	
	/**
	 * defines a hexadecimal background background color code
	 * 
	 * @param string $sBackgroundColor
	 * @return Style
	 */
	public function setBackgroundColor($sBackgroundColor) {
		$this->checkColor($sBackgroundColor);
		
		$this->sBackground = $sBackgroundColor;

		return $this;
	}

	/**
	 * return hexadecimal background background color code
	 * 
	 * @return null|string
	 */
	public function getBackgroundColor() {
		return $this->sBackground;
	}
	
	/**
	 * check if a color is a hexadecimal color
	 * 
	 * @param string $sColor
	 * @throws \InvalidArgumentException
	 */
	private function checkColor($sColor) {
		if (!preg_match('/^[A-F0-9]{6,6}$/', $sColor)) {
			throw new \InvalidArgumentException('Please provide a hexadecimal color code.');
		}
	}

#pragma mark - alignment

	/**
	 * defines horizontal alignment
	 *
	 * @param string $sAlignment
	 * @throws \InvalidArgumentException
	 * @return Style
	 */
	public function setHorizontalAlignment($sAlignment) {
		if (!in_array($sAlignment, array(self::ALIGN_CENTER, self::ALIGN_GENERAL, self::ALIGN_JUSTIFY, self::ALIGN_LEFT, self::ALIGN_RIGHT))) {
			throw new \InvalidArgumentException('Horizontal alignment allows only "center, general, justify, left, right".');
		}

		$this->sHorizontalAlignment = $sAlignment;

		return $this;
	}

	/**
	 * returns horizontal alignment
	 *
	 * @return null|string
	 */
	public function getHorizontalAlignment() {
		return $this->sHorizontalAlignment;
	}

	/**
	 * defines vertical alignment
	 *
	 * @param string $sAlignment
	 * @throws \InvalidArgumentException
	 * @return Style
	 */
	public function setVerticalAlignment($sAlignment) {
		if (!in_array($sAlignment, array(self::ALIGN_CENTER, self::ALIGN_TOP, self::ALIGN_JUSTIFY, self::ALIGN_BOTTOM))) {
			throw new \InvalidArgumentException('Horizontal alignment allows only "center, top,  justify, bottom".');
		}

		$this->sVerticalAlignment = $sAlignment;

		return $this;
	}

	/**
	 * returns vertical alignment
	 *
	 * @return null|string
	 */
	public function getVerticalAlignment() {
		return $this->sVerticalAlignment;
	}

#pragma mark - border

	/**
	 * define a border for cell. Border styles are stored by alignment,
	 * because alignment have to be unique.
	 *
	 * @param string $sStyle
	 * @param int $iAlignment
	 * @param string $sColor
	 *
	 * @return Style
	 * @throws \InvalidArgumentException
	 */
	public function setBorder($sStyle, $iAlignment, $sColor) {

		// allowed borders
		$aAllowedBorders = array(
			self::BORDER_NONE, self::BORDER_THIN, self::BORDER_MEDIUM, self::BORDER_THICK, self::BORDER_DOUBLE,
			self::BORDER_DOTTED, self::BORDER_DASH_DOT, self::BORDER_MEDIUM_DASH_DOT, self::BORDER_SLANT_DASH_DOT,
			self::BORDER_DASH_DOT_DOT, self::BORDER_MEDIUM_DASH_DOT_DOT, self::BORDER_DASHED, self::BORDER_MEDIUM_DASHED,
		);

		// check if border is allowed
		if (!in_array($sStyle, $aAllowedBorders)) {
			throw new \InvalidArgumentException(sprintf('Border style "%s" is not allowed. Please provide a valid style.', $sStyle));
		}

		// max border alignment bit
		$iMaxAlignment = self::BORDER_ALIGN_TOP | self::BORDER_ALIGN_RIGHT | self::BORDER_ALIGN_BOTTOM | self::BORDER_ALIGN_LEFT;

		// check if $iAlignment in that range
		if (0 == ($iMaxAlignment & $iAlignment)) {
			throw new \InvalidArgumentException('Please provide a valid border alignment style');
		}

		// check color
		$this->checkColor($sColor);

		// set border
		$this->aBorder[$iAlignment] = array(
			'style' => $sStyle,
			'color' => $sColor
		);

		return $this;
	}

	/**
	 * return border definition
	 *
	 * @return array
	 */
	public function getBorder() {
		return $this->aBorder;
	}
}