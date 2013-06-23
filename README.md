# Excellence

Excellence is a simple and lightweight implementation to write Excel files.
This library will only support _xlsx_ (OfficeOpenXml format) files.

__Library is currently under development__

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


### Excellence library

| Lines / Columns   | PHP Memory | Execution Time  | Filesize   |
| ----------------- | ---------- | --------------- | ---------- |
| 1000   / 20       | 4 MB       | ~ 1.7  seconds  |  102 KB    |
| 1000   / 100      | 15.5 MB    | ~ 10   seconds  |  483 KB    |
| 5000   / 20       | 17 MB      | ~ 8    seconds  |  513 KB    |
| 5000   / 100      | 76 MB      | ~ 56   seconds  |  2.4 KB    |
| 10000  / 20       | 33.25 MB   | ~ 19   seconds  |  1.0 MB    |
| 10000  / 100      | 151.75 MB  | ~ 116  seconds  |  4.7 MB    |
| 50000  / 20       | 164.25 MB  | ~ 200  seconds  |  5.1 MB    |
| 50000  / 100      | 757.75 MB  | ~ 1133 seconds  | 23.6 MB    |

The filesize against PHPExcel is a little bit higher, but i know how to fix this.
Currently i don't use string reference method in Excel. I will add this in further
development.

Maybe there is a way to speed it up more then this. I'm not happy about current
implementation of generation XML files. I will refactor this.

### PHPExcel library

| Lines / Columns   | PHP Memory | Execution Time | Filesize  |
| ----------------- | ---------- | -------------- | --------- |
| 1000   / 20       | 21.5   MB  | ~ 12  seconds  | 88  KB    |
| 1000   / 100      | 80.0   MB  | ~ 60  seconds  | 263 KB    |
| 5000   / 20       | 81.75  MB  | ~ 58  seconds  | 412 KB    |
| 5000   / 100      | 341.25 MB  | ~ 324 seconds  | 2.0 MB    |
| 10000  / 20       | 153.5  MB  | ~ 130 seconds  | 816 KB    |
| 10000  / 100      | 651.0  MB  | ~ 744 seconds  | 4 MB      |
| 50000  / 20       | 736 MB     | ~ 763 seconds  | 4.1 MB    |
| 50000  / 100      | > 2,7 GB   | > 70  minutes  | ?         |
