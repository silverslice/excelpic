What is it
============================================================

Easy tool for converting pictures in excel file to resizable with cells. It will be useful if you need
to collapse rows with images in xlsx-file creating with PHPExcel.

## Install

`composer require silverslice/excelpic`

## Example of usage

```php

use Silverslice\ExcelPic\Converter;

require __DIR__ . '/vendor/autoload.php';

$converter = new Converter();

// open xlsx file
$converter->open('test.xlsx');

    // convert all images to resizable
    ->convertImagesToResizable();

    // save xlsx document
    ->save('test.xlsx');

```

## Limitations

- Only first spreadsheet is processed.
- Picture should cover only one cell.