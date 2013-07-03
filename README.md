# Excellence

Excellence is a simple and lightweight implementation to write Excel files.
This library will only support _xlsx_ (OfficeOpenXml format) files.

__Library is currently under development__

## Roadmap

* Cell Merging
* Styles (border, color, font, size, padding ... )

## Performance Tests against PHPExcel library

Performance tests made on Mac OSX 10.9 (2.4 GHz Intel Core 2 Duo, 4 GB DDR3 Ram)
PHP 5.4 executing in terminal window. There where no styles added or something
else, this is just creating a simple lightweight Excel file.

Excel Table looks like this one:

| Rows  | A                 | B   | C   | D             | ... | T             |
| ----- | ----------------- | --- | --- | ------------- | --- | ------------- |
| 1     | test column value | 22  | 0.9 | =SUM(B1:C1)   | ... | =SUM(Q1:R1)   |
| 2     | test column value | 22  | 0.9 | =SUM(B2:C2)   | ... | =SUM(Q2:R2)   |
| ...   | ...               | ... | ... | ...           | ... | ...           |

### Excellence library vs PHPExcel

|                   |     | **Excellence** |                 |            |     | **PHPExcel** |                |            |
| ----------------- | --- | -------------: | --------------: | ---------: | --- |-----------:  | -------------: | ---------: |
| *Lines/Columns*   |     | *Memory*       | *Time*          | *Filesize* |     | *Memory*     | *Time*         | *Filesize* |
| 1000   / 20       |     | 5,25    MB     | ~ 0,7  seconds  |   87 KB    |     | 21.5   MB    | ~ 12  seconds  | 88  KB     |
| 1000   / 100      |     | 20.25   MB     | ~ 4,7  seconds  |  402 KB    |     | 80.0   MB    | ~ 60  seconds  | 263 KB     |
| 5000   / 20       |     | 21,50   MB     | ~ 4    seconds  |  416 KB    |     | 81.75  MB    | ~ 58  seconds  | 412 KB     |
| 5000   / 100      |     | 109,00  MB     | ~ 20   seconds  |  1,9 MB    |     | 341.25 MB    | ~ 324 seconds  | 2.0 MB     |
| 10000  / 20       |     | 42,00   MB     | ~ 7    seconds  |  826 KB    |     | 153.5  MB    | ~ 130 seconds  | 816 KB     |
| 10000  / 100      |     | 217,00  MB     | ~ 41   seconds  |  4,9 MB    |     | 651.0  MB    | ~ 744 seconds  | 4,0 MB     |
| 50000  / 20       |     | 229.25  MB     | ~ 37   seconds  |  4,0 MB    |     | 736 MB       | ~ 763 seconds  | 4,1 MB     |
| 50000  / 100      |     | 1022.75 MB     | ~ 213  seconds  | 19.3 MB    |     | > 2,7 GB     | > 70  minutes  | ?          |

## Usage

Excellence is based on [Delegation Pattern](http://www.blog.newventurewebsites.com/delegate-design-pattern-in-php/).
Excellence is providing interfaces that you have to implement. There are `WorkBookDelegate` and `DataDelegate` Interfaces.

### Workbook Class

The workbook class is to create a workbook. A workbook is created by an identifier string and a WorkbookDelegate instance.

```php
class Workbook implements \Excellence\Delegates\Workbook {
    // ...
}

$oWorkBookDelegate = new Workbook();
$oWorkbook = new \Excellence\Workbook('identifier', $oWorkBookDelegate);

```

### Sheet Class

A sheet object will set an identifier and a name for a sheet, that will displayed in Excel's Sheet bar.

```php
$oSheet = new \Excellence\Sheet('identifier', 'name of this sheet');
```

### WorkBookDelegate

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

### Data Delegate

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
			$this->aData['sheet1'][$iRow][0] = 'test column value';
			$this->aData['sheet1'][$iRow][1] = 22;
			$this->aData['sheet1'][$iRow][2] = 0.9;
			$this->aData['sheet1'][$iRow][3] = '=SUM(B1:C1)';
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

		// use this object as data delgate
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
		return count($this->aData[$oSheet->getIdentifier()][$iRow][$iColumn]);
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
		return count($this->aData[$oSheet->getIdentifier()][$iRow][$iColumn]);
	}
}

// create workbook and save it
$oWorkbook  =new Workbook('myWorkbook', new MyWorkbook());
$oWorkbook
	->create()
	->save('foobar.xmls')
;
```
