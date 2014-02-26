# StylableDelegate

The stylable delegte interface allows you to modify the appearance of a sheet.

#### Get standard styles

This method will load default style information and override Excellence default styles.

```php
/**
 * returns a Style definition that will be used as default for all cells.
 *
 * @param Workbook $oWorkbook
 * @param Sheet    $oSheet
 *
 * @return Style
 */
public function getStandardStyle(Workbook $oWorkbook, Sheet $oSheet);
```

#### Get style for column and row

This method will load specific styles for a specific cell identified by row number and
column number.

```php
/**
 * returns Style definition for a specific cell by given identifier
 *
 * @param Workbook $oWorkbook
 * @param Sheet    $oSheet
 * @param int      $iColumn
 * @param int      $iRow
 *
 * @return Style
 */
public function getStyleForColumnAndRow(Workbook $oWorkbook, Sheet $oSheet, $iColumn, $iRow);
```
