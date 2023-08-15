#!/usr/bin/env php
<?php
use Jgauthi\Component\Spreadsheet\{CsvFile, CsvUtils};

if (is_readable(__DIR__.'/../../../autoload.php')) {
    require_once __DIR__.'/../../../autoload.php';
} elseif (is_readable(__DIR__.'/../vendor/autoload.php')) {
    require_once __DIR__.'/../vendor/autoload.php';
} else {
    die('Autoloader not found');
}

$import = $argv[1];
$export = ((!empty($argv[2])) ? $argv[2] : dirname($import));

try {
	if (!is_readable($import)) {
		throw new ErrorException("Le fichier {$import} n'est pas accessible ou n'existe pas.");
    } elseif (!preg_match('#\.csv$#i', $import) || filesize($import) <= 10) {
        throw new ErrorException("Le fichier {$import} n'est pas un fichier csv valide.");
    }

    if (is_dir($export)) {
        if (!is_writable($export)) {
            throw new ErrorException("Le dossier d'export {$export} n'a pas les droits en écriture.");
        }

        $export .= '/'.str_replace('.csv', '.json', basename($import));
    } elseif (file_exists($export)) {
        unlink($export);
    }

	$csvContent = CsvUtils::csv_to_array(
	    $import,
        CsvFile::DELIMITER,
        CsvFile::ENCLOSURE,
        null,
        [ CsvUtils::EXTRACT_VALUE_TRIM, CsvUtils::EXTRACT_VALUE_EMPTY_IS_NULL, CsvUtils::EXTRACT_VALUE_NULL_STRING ]
    );

    $jsonContent = json_encode($csvContent, JSON_PRETTY_PRINT);
    if (file_put_contents($export, $jsonContent)) {
        echo "Fichier généré: {$export}".PHP_EOL;
    }

} catch (Throwable $exception) {
	die($exception->getMessage().PHP_EOL);
}


