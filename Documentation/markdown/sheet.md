# Sheet Class

A sheet object will set an identifier and a name for a sheet, that will displayed in Excel's Sheet bar.

```php
$oSheet = new \Excellence\Sheet('identifier', 'name of this sheet');
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
