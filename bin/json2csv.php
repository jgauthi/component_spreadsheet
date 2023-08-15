#!/usr/bin/env php
<?php
use Jgauthi\Component\Spreadsheet\CsvUtils;

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
	if (is_dir($export)) {
		if (!is_writable($export)) {
			throw new ErrorException("Le dossier d'export {$export} n'a pas les droits en écriture.");
		}

		$export .= '/'.str_replace('.json', '.csv', basename($import));
	} elseif (file_exists($export)) {
		unlink($export);
	}

	if (!is_readable($import)) {
		throw new ErrorException("Le fichier {$import} n'est pas accessible ou n'existe pas.");
	}

	$jsonContent = json_decode( file_get_contents($import), true );
	if (CsvUtils::generatecsv($export, $jsonContent)) {
		echo "Fichier généré: {$export}".PHP_EOL;
	}

} catch (Throwable $exception) {
	die($exception->getMessage().PHP_EOL);
}


