# Data Delegate

Data delagte instance will return only data for a specific sheet. There is nothing else to do.
It is your choice how to get the data you want to put into an Excel document. You can call a databse,
a rest service, a file, and array or whatever you want.

In all methods you have access to the current workbook and sheet. If you wan't to use same data
source for more than one workbook or sheet you can decide by `Workbook->getIdentifier()` and also
`Sheet->getIdentifier()`.


#### Number of rows in sheet

```php
/**
 * returns an integer value how many rows a sheet for workbook contains.
 *
 * @param Workbook $oWorkbook
 * @param Sheet $oSheet
 *
 * @return int
 */
public function numberOfRowsInSheet(Workbook $oWorkbook, Sheet $oSheet);
```

#### Number of columns in sheet

```php
/**
 * returns an integer value how many columns a sheet for workbook contains.
 *
 * @param Workbook $oWorkbook
 * @param Sheet $oSheet
 *
 * @return int
 */
public function numberOfColumnsInSheet(Workbook $oWorkbook, Sheet $oSheet);
```

#### Value for row and column

```php
/**
 * returns a value for sheet cell by given workbook, sheet, row number and
 * column number. possible values are integers, floats, doubles and strings.
 * An Excel function like "=SUM(A1:A4)" will also provided as string.
 *
 * @param Workbook $oWorkbook
 * @param Sheet $oSheet
 * @param integer $iRow
 * @param integer $iColumn
 *
 * @return string|float|double|int
 */
public function valueForRowAndColumn(Workbook $oWorkbook, Sheet $oSheet, $iRow, $iColumn);
```
