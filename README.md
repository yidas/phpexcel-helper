PHPExcel Helper
===============

Creating Excel with easy and artistic way based on PHPExcel

This library is a helper that encapsulate [PHPExcel](https://github.com/PHPOffice/PHPExcel/blob/1.8/Classes/PHPExcel/Worksheet.php) for simple usage.

---

DEMONSTRATION
-------------

```php
\PHPExcelHelper::newExcel()
    ->addRow(['ID', 'Name', 'Email'])
    ->addRows([
        ['1', 'Nick','myintaer@gmail.com'],
        ['2', 'Eric','eric@.....'],
    ])
    ->output('My Excel');
```

---

INSTALLATION
------------

Run Composer in your project:

    composer require yidas/phpexcel-helper
    
Then you could call it after Composer is loaded depended on your PHP framework:

```php
require __DIR__ . '/vendor/autoload.php';

\PHPExcelHelper::newExcel();
```
    
---

USAGE
-----

### Merge Cells

```php
\PHPExcelHelper::newExcel()
    ->addRows([
        [['value'=>'SN', 'row'=>2], ['value'=>'Language', 'col'=>2], ['value'=>'Block', 'row'=>2, 'col'=>2]],
        ['','English','繁體中文',['skip'=>2]],
    ])
    ->addRows([
        ['1', 'Computer','電腦','#15'],
        ['2', 'Phone','手機','#4','#62'],
    ])
    ->output('Merged Excel');
```

### PHPExcel & Sheet Object

```php
// Get a new PHPExcel object
$objPHPExcel = new \PHPExcel;
$objPHPExcel->getProperties()
    ->setCreator("Nick Tsai")
    ->setTitle("Office 2007 XLSX Document");
// Get the actived sheet object
$objPHPExcelSheet = $objPHPExcel->setActiveSheetIndex(0);
$objPHPExcelSheet->setTitle('Sheet');
$objPHPExcelSheet->setCellValue('A1', 'SN');
// Inject PHPExcel Object and Sheet Object to Helper
\PHPExcelHelper::newExcel($objPHPExcel)
    ->setSheet($objPHPExcelSheet)
    ->setRowOffset(1) // Point to 1nd row from 0
    ->addRows([
        ['1'],
        ['2'],
    ]);
    
\PHPExcelHelper::output();
```

```php
\PHPExcelHelper::newExcel()
    ->setSheet(0, 'Sheet')
    ->addRow(['SN']);
// Get the PHPExcel object created by Helper
$objPHPExcel = \PHPExcelHelper::getExcel();
$objPHPExcel->getProperties()
    ->setCreator("Nick Tsai")
    ->setTitle("Office 2007 XLSX Document");
// Get the actived sheet object created by Helper
$objPHPExcelSheet = \PHPExcelHelper::getSheet();
$objPHPExcelSheet->setCellValue('A2', '1');
$objPHPExcelSheet->setCellValue('A3', '2');

\PHPExcelHelper::output();
```

### Multiple Sheets

```php
\PHPExcelHelper::newExcel()
    ->setSheet(3, '4nd Sheet')
    ->addRow(['ID', 'Name'])
    ->addRows([
        ['1', 'Nick'],
    ]);
// Set another sheet object and switch to it    
\PHPExcelHelper::setSheet(1, '2nd Sheet')
    ->addRow(['SN', 'Title'])
    ->addRows([
        ['1', 'Foo'],
    ]);
    
\PHPExcelHelper::output('MultiSheets');
```

### Map of Coordinates & Ranges

```php
\PHPExcelHelper::newExcel()
    ->addRows([
        [
            ['value'=>'SN', 'row'=>2, 'key'=>'sn'], 
            ['value'=>'Language', 'col'=>2, 'key'=>'lang'], 
            ['value'=>'Block', 'row'=>2, 'col'=>2, 'key'=>'block'],
        ],
        [   
            '',
            ['value'=>'English', 'key'=>'lang-en'],
            ['value'=>'繁體中文', 'key'=>'lang-zh'],
            ['skip'=>2, 'key'=>'block-skip'],
        ],
    ])
    ->addRows([
        ['1', 'Computer','電腦','#15'],
        ['2', 'Phone','手機','#4','#62'],
    ]);
    // ->output('Merged Excel');  

print_r(\PHPExcelHelper::getCoordinateMap());
print_r(\PHPExcelHelper::getRangeMap());
echo \PHPExcelHelper::getCoordinateMap('sn');
```

The result could be:

```
Array
(
    [sn] => A1
    [lang] => B1
    [block] => D1
    [lang-en] => B2
    [lang-zh] => C2
    [block-skip] => D2
)
Array
(
    [sn] => A1:A2
    [lang] => B1:C1
    [block] => D1:E2
    [lang-en] => B2:B2
    [lang-zh] => C2:C2
    [block-skip] => D2:F2
)
A1
```

### Cells Format

* setWrapText(): Set to all cells by default
* setAutoSize(): Set to all cells(columns) by default

```php
\PHPExcelHelper::newExcel()
    ->addRow(['Title', 'Content'])
    ->addRows([
        ['Basic Plan', "*Interface\n*Search Tool"],
        ['Advanced Plan', "*Interface\n*Search Tool\n*Statistics"],
    ])
    ->setWrapText()
    // ->setWrapText('B2')
    ->setAutoSize()
    // ->setAutoSize('B')
    ->output('Formatted Excel');  
```
