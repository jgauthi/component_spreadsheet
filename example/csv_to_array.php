<?php
use Jgauthi\Component\Spreadsheet\CsvUtils;

// In this example, the vendor folder is located in "example/"
require_once __DIR__.'/vendor/autoload.php';

$file = __DIR__.'/asset/clients2.csv';
$csvContent = CsvUtils::csv_to_array($file, ';', '"');

var_dump($csvContent);
