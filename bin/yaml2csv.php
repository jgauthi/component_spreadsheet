#!/usr/bin/env php
<?php
use Jgauthi\Component\Spreadsheet\CsvUtils;
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
        throw new ErrorException("The YAML Component is not installed.");
    } elseif (!is_readable($import)) {
		throw new ErrorException("Le fichier {$import} n'est pas accessible ou n'existe pas.");
    } elseif (!preg_match('#\.ya?ml$#i', $import) || filesize($import) <= 10) {
        throw new ErrorException("Le fichier {$import} n'est pas un fichier yaml valide.");
    }

    if (is_dir($export)) {
        if (!is_writable($export)) {
            throw new ErrorException("Le dossier d'export {$export} n'a pas les droits en écriture.");
        }

        $export .= '/'.str_replace('.yaml', '.csv', basename($import));
    } elseif (file_exists($export)) {
        unlink($export);
    }

    $yamlContent = Yaml::parseFile($import);
    if (CsvUtils::generatecsv($export, $yamlContent)) {
        echo "Fichier généré: {$export}".PHP_EOL;
    }

} catch (Throwable $exception) {
	die($exception->getMessage().PHP_EOL);
}
