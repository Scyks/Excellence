
# Linking Example

The following example is a simple workbook that creates a sheet. The Class `MyWorkbook` will
handle sheet information. The first column is linked to http://www.google.de.


```php
use Excellence\Sheet;
use \Excellence\Workbook;
use \Excellence\Writer\Excel;
use \Excellence\Delegates\WorkbookDelegate;
use \Excellence\Delegates\DataDelegate;
use \Excellence\Delegates\LinkableDelegate;

class MyWorkbook implements WorkbookDelegate, DataDelegate, LinkableDelegate {

	private $aSheets = array();
	private $aData = array();

	public function __construct() {
		$this->aSheets[] = new Sheet('sheet1', 'Table 1');

		$this->aData['sheet2'][0][0] = 'google';
		$this->aData['sheet2'][0][1] = 22;
		$this->aData['sheet2'][0][2] = 0.9;
		$this->aData['sheet2'][0][3] = '=SUM(B1:C1)';
	}

	// workbook delegate

	/** simple count */
	public function numberOfSheetsInWorkbook(Workbook $oWorkbook) {
		return count($this->aSheets);
	}

	/** simple return */
	public function getSheetForWorkBook(Workbook $oWorkbook, $iSheetIndex) {
		return $this->aSheets[$iSheetIndex];
	}

	/** demonstrate usage of external and internal data source */
	public function dataSourceForWorkbookAndSheet(Workbook $oWorkbook, Sheet $oSheet) {
		return $this;
	}

	// Data Source delegate

	/** simple return  */
	public function numberOfRowsInSheet(Workbook $oWorkbook, Sheet $oSheet) {
		return count($this->aData[$oSheet->getIdentifier()]);
	}

	/** simple return  */
	public function numberOfColumnsInSheet(Workbook $oWorkbook, Sheet $oSheet) {
		return 4;
	}

	/** simple return  */
	public function valueForRowAndColumn(Workbook $oWorkbook, Sheet $oSheet, $iRow, $iColumn) {
		return $this->aData[$oSheet->getIdentifier()][$iRow][$iColumn];
	}

	/** returns true on first column **/
	public function hasLinkForColumnAndRow(Workbook $oWorkbook, Sheet $oSheet, $iRow, $iColumn) {
		return (1 == $iColumn);
	}

	/** simple return http://www.google.de*/
	public function getLinkForColumnAndRow(Workbook $oWorkbook, Sheet $oSheet, $iRow, $iColumn) {
		return 'http://www.google.de';
	}
}

// create workbook and save it
$oWorkbook = new Workbook('myWorkbook', new MyWorkbook());
$oWriter = new Excel($oWorkbook);

$oWriter->saveToFile('test.xlsx');

```

## Output

| A   | B     | C   | D   |
| --- | :---: | --- | --- |
| [google] | 22 | 0.9 | 22.9 |


[google]: <http://www.google.de> (test column value)