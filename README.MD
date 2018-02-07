# FatExcel
phpexcel / phpspreadsheet的功能非常全面，可对单元格进行各式各样的操作。
然而，它的缺点是，占用内存大，处理速度慢。在数据量小的时候可能没有影响，但是对于大数据量（例如10万+行），导入和导出花费的时间就会非常长。

为了解决这个问题，FatExcel基于其他的第三方开源类库，简单封装了大数据导入导出的方法。但是使用是有一定限制的，在后面会详细阐述。

# 安装
确保PHP版本在5.5以上。

推荐通过[Composer](https://getcomposer.org/)进行安装

Composer的安装请参考官方配置。

安装好Composer后，在你的项目中创建一个 composer.json 文件：
```json
{
    "require": {
        "tianhe1986/fatexcel": "*"
    }
}
```

然后在项目文件夹下执行:
```bash
$ composer install
```

之后添加autoloader:
```php
<?php
require "vendor/autoload.php";
```

# 导入
FatExcel类提供了`importToArray`方法，基于[spreadsheet-reader](https://github.com/nuovo/spreadsheet-reader)封装，将excel文件直接导入为二维数组，使用方法如下：
```php
use FatExcel\FatExcel;

$fatExcel = new FatExcel();

$data = $fatExcel->importToArray('/path/to/excel');
```

要使用该方法，对应的excel文件需要满足的条件是：
* 文件为xlsx格式(扩展名为.xlsx)
* 只有一个工作表(todo:支持多个工作表)
* 表格数据从A1开始，每行及每列单元格数量相同（没有单元格合并）


# 导出
与导入相对的，是从二维数组直接导出成excel，基于[PHP_XLSXWriter](https://github.com/mk-j/PHP_XLSXWriter)封装。使用方法如下：
```php
use FatExcel\FatExcel;

$fatExcel = new FatExcel();

$data = [
    [1, 2, 3, 4, 5],
    [6, 7, 8, 9, 10]
];

//导出到文件
$fatExcel->exportToFile($data, '/path/to/excel');

//直接下载
$excel->exportDownload($data, '666.xlsx');
```

对于单元格有格式要求的，可以通过表头数组进行设置，如下：
```php
use FatExcel\FatExcel;

$fatExcel = new FatExcel();

$header = [
    'created' => 'date',
    'product_id' => 'integer',
    'quantity' => '#,##0',
    'amount' => 'price',
    'total' => '0.00E+00',
    'code' => 'string',
    'tax' => '[$$-1009]#,##0.00;[RED]-[$$-1009]#,##0.00',
];

$data = [
    ['2015-01-01', 873, 1, '44.00', 99909, '12345678900000000', '=D2*0.05'],
    ['2015-01-12', 324, 2, '88.00', 10000, '12345678911111111','=D3*0.05'],
];

//导出到文件
$fatExcel->exportToFile($data, '/path/to/excel', $header);
```

对于`$header`数组的每项key和value，key将作为表头写入表格，此列的其他单元格将按照value定义的格式进行写入。

导出的注意事项为：
* 导出文件为xlsx格式。
* 不支持设置样式。
* 不支持单元格合并。
