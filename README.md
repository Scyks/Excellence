# Excellence

Excellence is a simple and lightweight implementation to write Excel files.
This library will only support _xlsx_ (OfficeOpenXml format) files.

__Library is currently under development__

## Performance Tests against PHPExcel library

Performance tests made on Mac OSX 10.9 (2.4 GHz Intel Core 2 Duo, 4 GB DDR3 Ram)
PHP 5.4 executing in terminal window. There where no styles added or something
else, this is just creating a simple lightweight Excel file.

Exel file looks like this one:

| Rows  | A                 | B  | C   | D             | .... | T             |
| ----- | ----------------- | -- | --- | ------------- | .... | ------------- |
| 1     | test column value | 22 | 0.9 | =SUM(B1:C1)   | .... | =SUM(Q1:R1)   |
| 2     | test column value | 22 | 0.9 | =SUM(B2:C2)   | .... | =SUM(Q2:R2)   |
| ...   | ...               | .. | ... | ...           | .... | ...           |

### Excellence library

| Lines / Columns   | PHP Memory | Execution Time | Filesize  |
| ----------------- | ---------- | -------------- | --------- |
| 1000   / 20       | 4 MB       | ~ 1.7 seconds  | 105 KB    |
| 5000   / 20       | 17 MB      | ~ 8   seconds  | 513 KB    |
| 10000  / 20       | 33.25 MB   | ~ 19  seconds  | 1 MB      |
| 50000  / 20       | 164.25 MB  | ~ 200 seconds  | 5.1 MB    |

### PHPExcel library

| Lines / Columns   | PHP Memory | Execution Time | Filesize  |
| ----------------- | ---------- | -------------- | --------- |
| 1000   / 20       | 21.5 MB    | ~ 34 seconds   | 88 KB     |
| 5000   / 20       | 81.75 MB   | ~ 85 seconds   | 411 KB    |
| 10000  / 20       | 152.5 MB   | ~ 154 seconds  | 816 KB    |
| 50000  / 20       | 736 MB     | ~ 800 seconds  | 4.1 MB    |