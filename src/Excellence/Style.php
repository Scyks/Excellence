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
	const BORDER_DOUBLE					= 'double';
	const BORDER_DOTTED					= 'dotted';
	const BORDER_DASH_DOT				= 'dashDot';
	const BORDER_MEDIUM_DASH_DOT		= 'mediumDashDot';
	const BORDER_SLANT_DASH_DOT			= 'slantDashDot';
	const BORDER_DASH_DOT_DOT			= 'dashDotDot';
	const BORDER_MEDIUM_DASH_DOT_DOT	= 'mediumDashDotDot';
	const BORDER_DASHED					= 'dashed';
	const BORDER_MEDIUM_DASHED			= 'mediumDashed';

	const BORDER_ALIGN_LEFT				= 1;
	const BORDER_ALIGN_RIGHT			= 2;
	const BORDER_ALIGN_TOP				= 4;
	const BORDER_ALIGN_BOTTOM			= 8;
	const BORDER_ALIGN_ALL				= 15;

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
	 * underline
	 * @var bool
	 */
	private $bUnderline = false;

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

	/**
	 * width of a cell
	 * 
	 * @var int
	 */
	private $fWidth = 0;

	/**
	 * height of a cell
	 * 
	 * @var int
	 */
	private $fHeight = 0;
	
	/**
	 * id of this styles
	 * @var null
	 */
	private $sId = null;

#pragma mark - identifier

	/**
	 * this method returns a string containing an id, by given params
	 * 
	 * @return string
	 */
	public function getId() {
		if (null == $this->sId) {
			$this->sId = md5(serialize($this));
		}

		return $this->sId;

	}

#pragma mark - font handling

	/**
	 * defines a font family, font size and color
	 *
	 * @param string $sFont
	 * @return Style
	 */
	public function setFont($sFont) {
		$this->sFont = $sFont;

		return $this;
	}

	/**
	 * checks if a font is defined
	 * @return bool
	 */
	public function hasFont() {
		return (null !== $this->sFont && is_string($this->sFont));
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
	 * checks if a font size is defined
	 * @return bool
	 */
	public function hasFontSize() {
		return (null !== $this->iFontSize && is_float($this->iFontSize));
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

	/**
	 * toggle function for marking font as underlined
	 * @param bool $bUnderline
	 * @return Style
	 */
	public function setUnderline($bUnderline = true) {
		$this->bUnderline = (bool) $bUnderline;

		return $this;
	}

	/**
	 * check if font is marked as underlined
	 *
	 * @return bool
	 */
	public function hasUnderline() {
		return $this->bUnderline;
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
		
		$this->sColor = strtoupper($sColor);

		return $this;
	}

	/**
	 * checks if a color is defined
	 * @return bool
	 */
	public function hasColor() {
		return (null !== $this->sColor && is_string($this->sColor));
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
		
		$this->sBackground = strtoupper($sBackgroundColor);

		return $this;
	}

	/**
	 * checks if a font is defined
	 * @return bool
	 */
	public function hasBackgroundColor() {
		return (null !== $this->sBackground && is_string($this->sBackground));
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
		if (!preg_match('/^[A-F0-9]{6,6}$/i', $sColor)) {
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
	 * checks if a horizontal alignment is defined
	 * @return bool
	 */
	public function hasHorizontalAlignment() {
		return (null !== $this->sHorizontalAlignment && is_string($this->sHorizontalAlignment));
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
			throw new \InvalidArgumentException('Vertical alignment allows only "center, top, justify, bottom".');
		}

		$this->sVerticalAlignment = $sAlignment;

		return $this;
	}

	/**
	 * checks if a vertical alignment is defined
	 * @return bool
	 */
	public function hasVerticalAlignment() {
		return (null !== $this->sVerticalAlignment && is_string($this->sVerticalAlignment));
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

		// check if $iAlignment in that range
		if (0 == (self::BORDER_ALIGN_ALL & $iAlignment)) {
			throw new \InvalidArgumentException('Please provide a valid border alignment style.');
		}

		// check color
		$this->checkColor($sColor);

		// set border

		for($i = 1; $i <= self::BORDER_ALIGN_ALL; $i = $i * 2) {
			if (($iAlignment & $i) == true) {
				$this->aBorder[$i] = array(
					'style' => $sStyle,
					'color' => strtoupper($sColor),
				);
			}
		}

		// this is very important, excel need border styles in left, right, top, bottom condition
		ksort($this->aBorder);

		return $this;
	}

	/**
	 * check if border is defined
	 *
	 * @return bool
	 */
	public function hasBorder() {
		return (!empty($this->aBorder));
	}

	/**
	 * return border definition
	 *
	 * @return array
	 */
	public function getBorder() {
		return $this->aBorder;
	}
	
#pragma mark - dimensions

	/**
	 * defines a width for a cell
	 *
	 * @param float $fWidth
	 * @return Style
	 */
	public function setWidth($fWidth) {
		
		$this->fWidth = (float) $fWidth;

		return $this;
	}

	/**
	 * checks if a width is defined
	 * @return bool
	 */
	public function hasWidth() {
		return (0 !== $this->fWidth);
	}

	/**
	 * return width for a cell
	 *
	 * @return null|float
	 */
	public function getWidth() {
		return $this->fWidth;
	}
	
	/**
	 * defines a height for a cell
	 *
	 * @param float $fHeight
	 * @return Style
	 */
	public function setHeight($fHeight) {
		
		$this->fHeight = (float) $fHeight;

		return $this;
	}

	/**
	 * checks if a height is defined
	 * @return bool
	 */
	public function hasHeight() {
		return (0 !== $this->fHeight);
	}

	/**
	 * return height for a cell
	 *
	 * @return null|float
	 */
	public function getHeight() {
		return $this->fHeight;
	}
}