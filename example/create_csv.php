<?php
use Jgauthi\Component\Spreadsheet\CsvFile;

// In this example, the vendor folder is located in "example/"
require_once __DIR__.'/vendor/autoload.php';

// CSV Generation
$result = json_decode(file_get_contents(__DIR__.'/asset/clients.json'), true);
if (isset($_GET['csv'])) {
    // $file = 'client.csv';
    $file = 'php://output';

    $csv = new CsvFile($file);
    $csv->writeTitle(array_keys($result[0]));

    foreach ($result as $data) {
        $csv->write($data);
    }

    $csv->close();
    // Alternative:
    // CsvUtils::generatecsv($file, $result);

    die();
}

?>
<p><a href="<?=$_SERVER['PHP_SELF']?>?csv">Get CSV File</a></p>
<p>Data:</p>
<?php var_dump($result);
