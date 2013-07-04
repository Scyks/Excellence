# Workbook Class

The workbook class is to create a workbook. A workbook is created by an identifier string and a WorkbookDelegate instance.

```php
class Workbook implements \Excellence\Delegates\Workbook {
    // ...
}

$oWorkBookDelegate = new Workbook();
$oWorkbook = new \Excellence\Workbook('identifier', $oWorkBookDelegate);

```

## Function reference

* Workbook **__construct** ( string *$sIdentifier* )

  Will return defined workbook delegate instance.

* string **getIdentifier** ( *void* )

  Will return defined identifier string.

* string **getCoordinatesByColumnAndRow** ( int *$iColumn*, int|null *$iRow = null* )

  Will return a coordinate string in Excel format. If *$iRow* is null, only the letter
  for a specific column will be returned. (A1)

* Workbook **create** ( *void* )

  This method create data structure for Excel document.

* Workbook **save** ( string *$sFilename*)

  will save an Excel document to file by given filename.
