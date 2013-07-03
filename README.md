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

| Lines / Columns   | PHP Memory  | Execution Time  | Filesize   |
| ----------------- | ----------- | --------------- | ---------- |
| 1000   / 20       | 5,25    MB  | ~ 0,7  seconds  |   87 KB    |
| 1000   / 100      | 20.25   MB  | ~ 4,7  seconds  |  402 KB    |
| 5000   / 20       | 21,50   MB  | ~ 4    seconds  |  416 KB    |
| 5000   / 100      | 109,00  MB  | ~ 20   seconds  |  1,9 MB    |
| 10000  / 20       | 42,00   MB  | ~ 7    seconds  |  826 KB    |
| 10000  / 100      | 217,00  MB  | ~ 41   seconds  |  4,9 MB    |
| 50000  / 20       | 229.25  MB  | ~ 37   seconds  |  4,0 MB    |
| 50000  / 100      | 1022.75 MB  | ~ 213  seconds  | 19.3 MB    |


### PHPExcel library

| Lines / Columns   | PHP Memory | Execution Time | Filesize  |
| ----------------- | ---------- | -------------- | --------- |
| 1000   / 20       | 21.5   MB  | ~ 12  seconds  | 88  KB    |
| 1000   / 100      | 80.0   MB  | ~ 60  seconds  | 263 KB    |
| 5000   / 20       | 81.75  MB  | ~ 58  seconds  | 412 KB    |
| 5000   / 100      | 341.25 MB  | ~ 324 seconds  | 2.0 MB    |
| 10000  / 20       | 153.5  MB  | ~ 130 seconds  | 816 KB    |
| 10000  / 100      | 651.0  MB  | ~ 744 seconds  | 4,0 MB    |
| 50000  / 20       | 736 MB     | ~ 763 seconds  | 4,1 MB    |
| 50000  / 100      | > 2,7 GB   | > 70  minutes  | ?         |
