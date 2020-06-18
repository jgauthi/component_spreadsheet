<?php
use Jgauthi\Component\Spreadsheet\CsvUtils;

require_once __DIR__.'/parse_csv.php';

//-- Database configuration (to complete) --------------------
$dbhost ??= 'localhost';
$dbuser ??= 'root';
$dbpass ??= 'root';
$dbname ??= 'dbname';
//------------------------------------------------------------

// Init PDO
$pdo = new PDO("mysql:dbname={$dbname};host={$dbhost}", $dbuser, $dbpass, [
    PDO::MYSQL_ATTR_INIT_COMMAND    => 'SET NAMES utf8 COLLATE utf8_unicode_ci',
    PDO::ATTR_DEFAULT_FETCH_MODE    => PDO::FETCH_ASSOC,
]);

// [Alternative] Get PDO instance with Library
// $pdo = $entityManager->getConnection()->getWrappedConnection(); // Symfony Entity Manager (doctrine)
// $pdo = $dbal->getWrappedConnection(); // Doctrine DBAL
// $pdo = $db->getPdoVar();  // github.com/jgauthi/indieteq-php-my-sql-pdo-database-class


// Convert CSV content to mysql tables ($csvContent is set by CsvUtils::parsecsv on previous example)
$table = 'clients_import_csv';
$createTable = true; // Create table (optional) + insert data (on existing table if false)
$processCreateTable = CsvUtils::content_to_database(
    $pdo,
    $csvContent['title'],
    $csvContent['content'],
    $table,
    $createTable
);

if (!$processCreateTable) {
    die("<p>Table {$table}: Echec creation.</p>");
}
echo "<p>Table {$table}: Created with success.</p>";

$clientContent = $pdo->query('SELECT * FROM '.$table)->fetchAll();
var_dump($clientContent);
