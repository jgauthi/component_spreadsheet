<?php
use Jgauthi\Component\Spreadsheet\ExcelXmlFile;

// In this example, the vendor folder is located in "example/"
require_once __DIR__.'/vendor/autoload.php';

// CSV Generation
$result = json_decode(file_get_contents(__DIR__.'/asset/clients.json'), true);
if (isset($_GET['xml'])) {
    // $file = 'client.xlsx.xml';
    $file = 'php://output';

    $xml = new ExcelXmlFile($file);
    $xml
        ->header('John doe', new DateTimeImmutable('2020-06-17 17:00'))
        ->writeTitle(array_keys($result[0]));

    foreach ($result as $line) {
        $xml->lineOpen();
        foreach ($line as $col) {
            $xml->writeCol($col);
        }

        $xml->lineClose();
    }

    $xml->footer();

    // Alternative
    // $xml = new ExcelXmlFile($file);
    // $xml->array_to_xml($result, 'John Doe');
    die();
}

?>
<p><a href="<?=$_SERVER['PHP_SELF']?>?xml">Get Excel XML File</a></p>
<p>Data:</p>
<?php var_dump($result);
