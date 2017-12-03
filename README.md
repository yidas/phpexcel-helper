PHPExcel Helper
===============

Creating Excel with easy and artistic way based on PHPExcel

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

Run Composer in your Codeigniter project:

    composer require yidas/phpexcel-helper
    
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
$objPHPExcel = new \PHPExcel;
$objPHPExcel->getProperties()
    ->setCreator("Nick Tsai")
    ->setTitle("Office 2007 XLSX Document");
$objPHPExcelSheet = $objPHPExcel->setActiveSheetIndex(0);
$objPHPExcelSheet->setTitle('Sheet');
$objPHPExcelSheet->setCellValue('A1', 'SN');
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
$objPHPExcel = \PHPExcelHelper::getExcel();
$objPHPExcel->getProperties()
    ->setCreator("Nick Tsai")
    ->setTitle("Office 2007 XLSX Document");
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
\PHPExcelHelper::setSheet(1, '2nd Sheet')
    ->addRow(['SN', 'Title'])
    ->addRows([
        ['1', 'Foo'],
    ]);
\PHPExcelHelper::output('MultiSheets');
```

