# WorkBookDelegate

The workbook delegte interface allows you to create sheets in a workbook and return a datasource for each sheet.

#### Number of Sheets in Workbook

This method is used to define how many sheets availble for a workbook.

```php
/**
 * returns integer value of how many sheets this workbook
 * will contain.
 *
 * @param Workbook $oWorkbook
 *
 * @return int
 */
public function numberOfSheetsInWorkbook(Workbook $oWorkbook);
```

#### Get sheet for Workbook

This method is used to create Sheets for a workbook and return it.

```php
/**
 * Return a Sheet for Workbook by given sheet index. If there
 * three sheets (numberOfSheetsInWorkbook) available, this
 * method will called three times by provide sheet index 0, 1
 * and 2.
 *
 * @param Workbook $oWorkbook
 * @param integer $iSheetIndex
 *
 * @return Sheet
 */
public function getSheetForWorkBook(Workbook $oWorkbook, $iSheetIndex);
```

#### data source for workbok and sheet

This method allows you to define and return a data source instance for each sheet. When
you have 3 Sheets available you can return 3 different data sources or even the same.
You can dicide you your code works.

```php
/**
 * Return a data source instance for given workbook and sheet.
 * This method will be called n times. N is the number retrieved
 * by numberOfSheetsInWorkbook.
 *
 * @param Workbook $oWorkbook
 * @param Sheet $oSheet
 * @return mixed
 */
public function dataSourceForWorkbookAndSheet(Workbook $oWorkbook, Sheet $oSheet);
```
