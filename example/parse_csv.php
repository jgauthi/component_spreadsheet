<?php
use Jgauthi\Component\Spreadsheet\CsvFile;
use Jgauthi\Component\Spreadsheet\CsvUtils;

// In this example, the vendor folder is located in "example/"
require_once __DIR__.'/vendor/autoload.php';

$file = __DIR__.'/asset/clients2.csv';
$csvContent = CsvUtils::parsecsv($file,
    true,
    CsvFile::DELIMITER,
    [ CsvUtils::EXTRACT_VALUE_TRIM, CsvUtils::EXTRACT_VALUE_NULL_STRING, CsvUtils::EXTRACT_VALUE_EMPTY_IS_NULL ]
);

var_dump($csvContent);
