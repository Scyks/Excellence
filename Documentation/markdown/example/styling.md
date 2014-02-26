
# Styling Example

The following example is a simple workbook that creates a sheet. The Class `MyWorkbook` will
handle sheet information. All Columns are styled manually to generate a nice looking sheet.


```php
use Excellence\Sheet;
use \Excellence\Workbook;
use \Excellence\Writer\Excel;
use \Excellence\Delegates\WorkbookDelegate;
use \Excellence\Delegates\DataDelegate;
use \Excellence\Delegates\StylableDelegate;

class MyWorkbook implements WorkbookDelegate, DataDelegate, StylableDelegate {

	private $aSheets = array();
	private $aData = array();

	public function __construct() {
		$this->aSheets[] = new Sheet('sheet1', 'Table 1');
		$this->aSheets[0]->setFirstRowAsFixed(true);

		// header
		$this->aData['sheet1'][0][0] = 'some text';
		$this->aData['sheet1'][0][1] = 'some integer';
		$this->aData['sheet1'][0][2] = 'some float';
		$this->aData['sheet1'][0][3] = 'some calculation';

		// values
		for ($iRow = 1; $iRow < $iRows; $iRow++) {
			$this->aData['sheet1'][$iRow][0] = 'Hello world';
			$this->aData['sheet1'][$iRow][1] = 22;
			$this->aData['sheet1'][$iRow][2] = 0.9;
			$this->aData['sheet1'][$iRow][3] = '=SUM(B' . ($iRow+1) . ':'C' . ($iRow+1) . ')';
		}

		$this->oStyle = new Style();
		$this->oStyle
			->setBackgroundColor('FCFAE4')
			->setFont('Verdana')
			->setFontSize(12)
			->setColor('000000')
			->setBorder(Style::BORDER_THIN, Style::BORDER_ALIGN_ALL, '333333')
			->setVerticalAlignment(Style::ALIGN_CENTER)
			->setHorizontalAlignment(Style::ALIGN_CENTER)
			->setHeight(30)
		;


		// header Styles
		$this->oHeader = clone $this->oStyle;
		$this->oHeader
			->setBold()
			->setBackgroundColor('FBF072')
			->setHeight(50)
		;

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

	/** returns standard Style */
	public function getStandardStyle(Workbook $oWorkbook, Sheet $oSheet) {
		$oStyle = new Style();
		$oStyle
			->setFont('Verdana')
			->setFontSize(12)
			->setColor('000000')
		;

		return $oStyle;
	}

    /** returns Style definition for a cell */
	public function getStyleForColumnAndRow(Workbook $oWorkbook, Sheet $oSheet, $iColumn, $iRow) {
		if (0 == $iRow) {
			return clone $this->oHeader;
		}


		// different colors for every second row
		if ($iRow %2 == 1) {
			$this->oStyle->setBackgroundColor('FCFAE4');
			return clone $this->oStyle;
		} else {
			$this->oStyle->setBackgroundColor('FBF6BE');
			return clone $this->oStyle;
		}

	}
}

// create workbook and save it
$oWorkbook = new Workbook('myWorkbook', new MyWorkbook());
$oWriter = new Excel($oWorkbook);

$oWriter->saveToFile('test.xlsx');

```