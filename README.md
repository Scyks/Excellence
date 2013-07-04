# Excellence

Excellence is a simple and lightweight implementation to write Excel files.
This library will only support _xlsx_ (OfficeOpenXml format) files.

__Library is currently under development__

[Documentation](Documentation/markdown/index.md)

## Roadmap

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

|                   |     | **Excellence** |           |            |     | **PHPExcel** |           |            |
| ----------------- | --- | -------------: | --------: | ---------: | --- |-----------:  | --------: | ---------: |
| *Lines/Columns*   |     | *Memory*       | *Time*    | *Filesize* |     | *Memory*     | *Time*    | *Filesize* |
| 1000   / 20       |     | 4,75    MB     | ~ 0,9 sec |   85  KB   |     | 21.5   MB    | ~ 12  sec | 86   KB    |
| 5000   / 20       |     | 21,00   MB     | ~ 4,3 sec |  406  KB   |     | 81.75  MB    | ~ 55  sec | 402  KB    |
| 10000  / 20       |     | 41,00   MB     | ~ 9   sec |  806  KB   |     | 153.5  MB    | ~ 114 sec | 796  KB    |
| 50000  / 20       |     | 204.00  MB     | ~ 44  sec |  3,8  MB   |     | 736 MB       | ~ 763 sec | 4,1  MB    |
| 1000   / 100      |     | 19.50   MB     | ~ 5,0 sec |  393  KB   |     | 77.75  MB    | ~ 59  sec | 388  KB    |
| 5000   / 100      |     | 95,50   MB     | ~ 25 sec  |  1,86 MB   |     | 367.75 MB    | ~ 339 sec | 1,89 MB    |
| 10000  / 100      |     | 190,75  MB     | ~ 51  sec |  3,7  MB   |     | 651.0  MB    | ~ 744 sec | 4,0  MB    |
| 50000  / 100      |     | 957.75  MB     | ~ 266 sec | 18.4  MB   |     | > 2,7 GB     | > 70  min | ?          |


