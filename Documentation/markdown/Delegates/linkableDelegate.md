# LinkableDelegate

The Linkable delegate interface allows you to link a cell value to an external source like
a web site.

#### Has link for Column and Row

This method is used to check if a specific cell has a hyperlink or not.

```php
/**
 * this method will return true when a specific column in a specific
 * row has an hyperlink. If this column has any, the method getLinkForColumnAndRow
 * will called.
 *
 * @param Workbook $oWorkbook
 * @param Sheet    $oSheet
 * @param int      $iRow
 * @param int      $iColumn
 *
 * @return boolean
 */
public function hasLinkForColumnAndRow(Workbook $oWorkbook, Sheet $oSheet, $iRow, $iColumn);
```

#### Get link for column and row

This method is used to retrieve the link for a cell. It will only be called, when
`hasLinkForColumnAndRow` returns true.

```php
/**
 * this method will return a hyperlink (url) for a specific column in a specific
 * row. This method is only been called, when getLinkForColumnAndRow returns
 * true.
 *
 * @param Workbook $oWorkbook
 * @param Sheet    $oSheet
 * @param int      $iRow
 * @param int      $iColumn
 *
 * @return string
 */
public function getLinkForColumnAndRow(Workbook $oWorkbook, Sheet $oSheet, $iRow, $iColumn);
```