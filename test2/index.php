<?php

require __DIR__ . '/vendor/autoload.php';

use IwanBazz\ExcelValidator;

echo "Validate Type_A\n";
$excelValidator = new ExcelValidator('./input_file/Type_A.xlsx');
$excelValidator->validate('A', 0);
$excelValidator->printOut();

echo "Validate Type_B\n";
$excelValidator = new ExcelValidator('./input_file/Type_B.xlsx');
$excelValidator->validate('B', 0);
$excelValidator->printOut();

echo "Validate Type_C\n";
$excelValidator = new ExcelValidator('./input_file/Type_C.xls');
$excelValidator->validate('C', 0);
$excelValidator->printOut();
