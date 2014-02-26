# Style Class

A style object allows to modify most used style information in an excel sheet as font, font size,
bold, italic, underline, color, background color, vertical and horizontal alignments and borders.

```php
$oStyle = new \Excellence\Style();
$oStyle
	->setBackgroundColor('FCFAE4')
	->setFont('Verdana')
	->setFontSize(12)
	->setColor('000000')
	->setBorder(Style::BORDER_THIN, Style::BORDER_ALIGN_ALL, '333333')
	->setVerticalAlignment(Style::ALIGN_CENTER)
	->setHorizontalAlignment(Style::ALIGN_CENTER)
	->setHeight(30)
;
```

## Function reference

* string **getId** (  )

  this method returns a string containing an id, by given params

* Style **setFont** ( *string* )

  defines a font family, font size and color

* bool **hasFont** (  )

  checks if a font is defined

* string **getFont** (  )

  return font family

* Style **setFontSize** ( *integer* )

  defines a font size - will casted to float value

* bool **hasFontSize** ( )

  checks if a font size is defined

* float **getFontSize** ( )

  return defined font size

* Style **setBold** ( *bool* )

  toggle function for marking font as bold

* bool **isBold** (  )

  check if font is marked as bold

* Style **setItalic** ( *bool* )

  toggle function for marking font as italic

* bool **isItalic** (  )

  check if font is marked as bold

* Style **setUnderline** ( *bool* )

  toggle function for marking font as underlined

* bool **hasUnderline** (  )

  check if font is marked as underlined

* Style **setColor** ( *string* )

  defines a hexadecimal color code

* bool **hasColor** (  )

  checks if a color is defined

* string **getColor** (  )

  return hexadecimal color code

* Style **setBackgroundColor** ( *string* )

  defines a hexadecimal background background color code

* bool **hasBackgroundColor** (  )

  checks if a font is defined

* string **getBackgroundColor** (  )

  return hexadecimal background background color code

* Style **setHorizontalAlignment** ( *string* )

  defines horizontal alignment

* bool **hasHorizontalAlignment** (  )

  checks if a horizontal alignment is defined

* string **getHorizontalAlignment** (  )

  returns horizontal alignment

* Style **setVerticalAlignment** ( *string* )

  defines vertical alignment

* bool **hasVerticalAlignment** (  )

  checks if a vertical alignment is defined

* string **getVerticalAlignment** (  )

  returns vertical alignment

* Style **setBorder** ( *string*, *integer*, *string* )

  define a border for cell. Border styles are stored by alignment, because alignment have to be unique.

* bool **hasBorder** (  )

  check if border is defined

* array **getBorder** (  )

  return border definition

* Style **setWidth** ( *float* )

  defines a width for a cell

* bool **hasWidth** (  )

  checks if a width is defined

* float **getWidth** (  )

  return width for a cell

* Style **setHeight** ( *float* )

  defines a height for a cell

* bool **hasHeight** (  )

  checks if a height is defined

* float **getHeight** (  )

  return height for a cell

