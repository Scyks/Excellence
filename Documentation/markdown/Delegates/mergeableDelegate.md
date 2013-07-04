# Mergable Delegate

Exellence will call your delegate class only if you implement `MergeableDelegate` interface. This interface method will called at every column and row. If the current coords souldn't be a merged cell just return false, true or null. Excellence will only merge a cell if this method return a string and contains a valid range. If a merge range is not valid, Excellence will throw an invalid argument exception.

```php
/**
 * Returns a string merge definition as string. A merge string will
 * contain two cell coordinates separated by a colon. Workbook's API
 * provide a method to retrieve a cell coordinate by column and row.
 *
 * - Workbook::getCoordinatesByColumnAndRow(int $iColumn, int $iRow)
 *
 * Format examples:
 * - A1:B1
 * - A1:B2
 *
 * @param Workbook $oWorkbook
 * @param int $iColumn
 * @param int $iRow
 *
 * @return string
 */
public function mergeByColumnAndRow(Workbook $oWorkbook, $iColumn, $iRow);
```