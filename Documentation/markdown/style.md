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

* Sheet **__construct** ( string *$sIdentifier*, string *$sName = null* )

  Will create a sheet instance with an identifier. You can also set a name
  for this sheet that will displayed in sheet tab bar in Excel.

* string **getIdentifier** ( *void* )

  Return sheet identifier to identify a workbook. This should be unique
  in an application.

* bool **hasName** ( *void* )

  Return true if this sheet has a name, otherise this method return false.

* string **getName** ( *void* )

  Return defined sheet name

* Sheet **setFirstRowAsFixed** ( *bool* )

  Will mark the first line as fixed, that i can't be scrolled.

* bool **isFirstRowFixed** ( )

  Check method if first row is marked as fixed
