# Excel Writer Class

This writer is an implementation for Excel. I decided to decouple the writing process from
the main classes that it become possible to genrate a writer that produces HTML, PDF or something else.
Every writer has to extend from `\Excellence\Writer\AbstractWriter` that provides main functionality.


```php
$oWorkBookDelegate = new Workbook();
$oWorkbook = new \Excellence\Workbook('identifier', $oWorkBookDelegate);

$oWriter = new \Excellence\Writer\Excel($oWorkbook);
$oWriter->saveToFile('/tmp/my_excel_document.xlsx');
```

## Function reference

* Sheet **__construct** ( Workbook *$oWorkbook*)

  Will create a Writer instance with a given workbook.

* string **getWorkbook** ( *void* )

  Getter implementation for give workbook class.

* WorkbookDelegate **getDelegate** ( *void* )

  Will return an instance of WorkbookDelegate bx calling Workbook class.

* bool **saveToFile** ( *string* filename )

  This is the main functionality of this class, that creates an excel document and save it to
  a file by given destination. In this method all delegate methods are called to generate the
  excel file.
