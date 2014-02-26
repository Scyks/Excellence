# Cell merging

To merge cells, you have to implement `Excellence\Delegates\MergeableDelegate` interface
in your `DataDelegate` class. The mergable delegate interface includes a method to
enable cell merging. See documentation of [Mergeeable Delegate](Delegates/mergeableDelegate.md).

## Merge range format

A merge range contains two coordinates separated by a colon. Coordinates is a letter that
specified the column followed by a row number. When you want to merge cell B1 and C1 the
format ist `B1:C1`. You can also merge more than one row. You can merge a cell range that
contains cell B1, B2, C1, C2 by format `B1:C2`. Of course you can also merge A1 and A2 to
one cell by format `A1:A2`.

|     | A   | B   | C   | D   |
| --- | :---: | --- | --- | --- |
| 1   | A1  | B1  | C1  | D1  |
| 2   | A2  | B2  | C2  | D2  |

## Merge examples
<table>
	<tr>
		<th align="center">&nbsp;</th>
		<th align="center">A</th>
		<th align="center">B</th>
		<th align="center">C</th>
		<th align="center">D</th>
		<th align="center" rowspan="3">&nbsp;</th>

		<th align="center">&nbsp;</th>
		<th align="center">A</th>
		<th align="center">B</th>
		<th align="center">C</th>
		<th align="center">D</th>
		<th align="center" rowspan="3">&nbsp;</th>

		<th align="center">&nbsp;</th>
		<th align="center">A</th>
		<th align="center">B</th>
		<th align="center">C</th>
		<th align="center">D</th>
	</tr>
	<tr>
		<th align="center">1</th>
		<td align="center">A1</td>
		<td align="center" colspan="2">B1:C1</td>
		<td align="center">D1</td>

		<th align="center">1</th>
		<td align="center">A1</td>
		<td align="center" colspan="2" rowspan=2">B1:C2</td>
		<td align="center">D1</td>

		<th align="center">1</th>
		<td align="center" rowspan="2">A1:A2</td>
		<td align="center">B1</td>
		<td align="center">C1</td>
		<td align="center">D1</td>
	</tr>
	<tr>
		<th align="center">2</th>
		<td align="center">A2</td>
		<td align="center">B2</td>
		<td align="center">C2</td>
		<td align="center">D2</td>

		<th align="center">2</th>
		<td align="center">A2</td>
		<td align="center">D2</td>

		<th align="center">2</th>
		<td align="center">B2</td>
		<td align="center">C2</td>
		<td align="center">D2</td>
	</tr>
</table>

## Example implementation

```php
use Excellence\Delegates\MergeableDelegate;
use Excellence\Delegates\DataDelegate;
use \Excellence\Workbook;
use \Excellence\Sheet;

class MergeableDataSource extends DataSource implements MergeableDelegate, DataDelegate {
	/**
	 * construction - load data
	 */
	public function __construct() {
		$this->aSheets[] = new Sheet('sheet1', 'Sheet 1');

		$this->aData['sheet1'][0][] = 'row1col1';	// A0
		$this->aData['sheet1'][0][] = 'B1:C1';		// B0
		$this->aData['sheet1'][1][] = 'row2col1';	// A1
		$this->aData['sheet1'][1][] = 'B2:C3';		// B1
		$this->aData['sheet1'][2][] = 'row3col1';	// A2
		$this->aData['sheet1'][3][] = 'A4:A5';		// A3

	}

	// implement DataDelegate source code here ....

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
	public function mergeByColumnAndRow(Workbook $oWorkbook, $iColumn, $iRow) {
		if (1 == $iColumn && 0 == $iRow) {
			return 'B1:C1';
		} elseif (1 == $iColumn && 1 == $iRow) {
			return 'B2:C3';
		} elseif (0 == $iColumn && 3 == $iRow) {
			return 'A3:A4';
		}

		return null;
	}
```