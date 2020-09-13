#!/usr/bin/env php
<?php
use Jgauthi\Component\Spreadsheet\{CsvFile, CsvUtils};
use Symfony\Component\Yaml\Yaml;

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
    if (!class_exists('Symfony\Component\Yaml\Yaml')) {
        throw new Exception("The YAML Component is not installed.");
    } elseif (!is_readable($import)) {
		throw new Exception("Le fichier {$import} n'est pas accessible ou n'existe pas.");
    } elseif (!preg_match('#\.csv$#i', $import) || filesize($import) <= 10) {
        throw new Exception("Le fichier {$import} n'est pas un fichier csv valide.");
    }

    if (is_dir($export)) {
        if (!is_writable($export)) {
            throw new Exception("Le dossier d'export {$export} n'a pas les droits en écriture.");
        }

        $export .= '/'.str_replace('.csv', '.yaml', basename($import));
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

    $yamlContent = Yaml::dump($csvContent, 4, 2, Yaml::DUMP_OBJECT);
    if (file_put_contents($export, $yamlContent)) {
        echo "Fichier généré: {$export}\n";
    }

} catch (Exception $exception) {
	die($exception->getMessage()."\n");
}
