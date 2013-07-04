
# Example

The following example is a simple workbook that creates 2 sheets. The Class `MyWorkbook` will
handle sheet information and data source for second sheet. For the first sheet there is
and external data source that holds same content as sheet 2. This is just to demonstrate
how you can use it.


```php
use Excellence\Sheet;
use \Excellence\Workbook;
use \Excellence\Delegates\WorkbookDelegate;
use \Excellence\Delegates\DataDelegate;

class MyWorkbook implements WorkbookDelegate, DataDelegate {

	private $aSheets = array();
	private $aData = array();

	public function __construct() {
		$this->aSheets[] = new Sheet('sheet1', 'Table 1');
		$this->aSheets[] = new Sheet('sheet2', 'Table 2');

		for ($iRow = 0; $iRow < 10; $iRow++) {
			$this->aData['sheet2'][$iRow][0] = 'test column value';
			$this->aData['sheet2'][$iRow][1] = 22;
			$this->aData['sheet2'][$iRow][2] = 0.9;
			$this->aData['sheet2'][$iRow][3] = '=SUM(B1:C1)';
		}
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

	/** demonstarte usage of external and internal data source */
	public function dataSourceForWorkbookAndSheet(Workbook $oWorkbook, Sheet $oSheet) {
		if ('sheet1' == $oSheet->getIdentifier()) {
			return new MyDataSource();
		}

		// use this object as data delegate
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

}

/** external standalone data source */
class myDataSource implements DataDelegate {

	// Data Source delegate

	private $aData = array();

	/** create data */
	public function __construct() {

		for ($iRow = 0; $iRow < 10; $iRow++) {
			$this->aData['sheet1'][$iRow][0] = 'test column value';
			$this->aData['sheet1'][$iRow][1] = 22;
			$this->aData['sheet1'][$iRow][2] = 0.9;
			$this->aData['sheet1'][$iRow][3] = '=SUM(B1:C1)';
		}
	}

	/** simple return  */
	public function numberOfRowsInSheet(Workbook $oWorkbook, Sheet $oSheet) {
		return count($this->aData[$oSheet->getIdentifier()]);
	}

	/** simple return */
	public function numberOfColumnsInSheet(Workbook $oWorkbook, Sheet $oSheet) {
		return 4;
	}

	/** simple return */
	public function valueForRowAndColumn(Workbook $oWorkbook, Sheet $oSheet, $iRow, $iColumn) {
		return $this->aData[$oSheet->getIdentifier()][$iRow][$iColumn];
	}
}

// create workbook and save it
$oWorkbook  =new Workbook('myWorkbook', new MyWorkbook());
$oWorkbook
	->create()
	->save('foobar.xmls')
;
```
