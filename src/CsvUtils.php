<?php
/*******************************************************************************
  * @name: CsvUtils
  * @note: Utility functions to create / read / generate CSV
  * @author: Jgauthi, created at [23oct2018], url: <github.com/jgauthi/component_spreadsheet>
  * @version: 1.1
  * @Requirements:
        - PHP version >= 8.2+ (http://php.net)

*******************************************************************************/

namespace Jgauthi\Component\Spreadsheet;

use InvalidArgumentException;
use SplFileObject;

class CsvUtils
{
    public const EXTRACT_VALUE_TRIM = 'trim';
    public const EXTRACT_VALUE_EMPTY_IS_NULL = 'null_empty';
    public const EXTRACT_VALUE_NULL_STRING = 'null_string';
    public const EXTRACT_VALUE_DEFAULT_MODE = [
        self::EXTRACT_VALUE_TRIM,
        self::EXTRACT_VALUE_EMPTY_IS_NULL,
    ];

    public static function parsecsv(
        string $filename,
        bool $first_line_title = true,
        string $delimiter = CsvFile::DELIMITER,
        array $extractValueMode = self::EXTRACT_VALUE_DEFAULT_MODE,
    ): array {
        $csv = [
            'filename' => basename($filename),
            'title' => [],
            'content' => [],
        ];

        if (!file_exists($filename)) {
            return $csv;
        }

        $row = $row_content = 0;

        $file = new SplFileObject($filename);
        $file->setFlags(
            SplFileObject::READ_CSV |
                SplFileObject::READ_AHEAD |
                SplFileObject::SKIP_EMPTY |
                SplFileObject::DROP_NEW_LINE
        );
        $file->setCsvControl($delimiter);

        foreach ($file as $data) {
            if (0 === $row && $first_line_title) {
                $csv['title'] = array_map('trim', $data);
            } else {
                $csv['content'][$row] = [];

                foreach ($data as $var) {
                    $var = self::csvContentExtractField($var, $extractValueMode);
                    $csv['content'][$row][] = $var;

                    $row_content += mb_strlen($var ?? '');
                }

                // Vérifier que la ligne inscrite n'est pas vide
                if (0 === $row_content) {
                    unset($csv['content'][$row][count($csv['content'][$row]) - 1]);
                    --$row;
                } else {
                    $row_content = 0;
                }
            }

            ++$row;
        }

        return $csv;
    }

    public static function csv_to_array(
        string $file,
        string $delimiter = CsvFile::DELIMITER,
        string $enclosure = CsvFile::ENCLOSURE,
        ?string $key_line = null,
        array $extractValueMode = self::EXTRACT_VALUE_DEFAULT_MODE,
    ): array {
        if (!is_readable($file)) {
            throw new InvalidArgumentException("The file {$file} doesn't exists or not readable.");
        }

        $titles = $content = [];

        if (false === ($handle = fopen($file, 'r'))) {
            return $content;
        }

        // Détection et suppression du BOM UTF8
        $bom = "\xEF\xBB\xBF";
        $firstLine = fgets($handle);

        if (strpos($firstLine, $bom) === 0) {
            $firstLine = substr($firstLine, 3);
        }

        // Analyser la première ligne pour les titres
        $data = str_getcsv($firstLine, $delimiter, $enclosure, $escape);
        foreach ($data as $index => $var) {
            $titles[$index] = trim($var);
        }

        // Intégration du contenu
        while (false !== ($data = fgetcsv($handle, 3000, $delimiter, $enclosure))) {
            // Récupérer le contenu
            $line = [];
            $row_content = 0;

            foreach ($data as $index => $var) {
                if (!isset($titles[$index])) {
                    continue;
                }

                $var = self::forceUtf8($var);
                $line[$titles[$index]] = self::csvContentExtractField($var, $extractValueMode);

                $row_content += mb_strlen($var);
            }

            // Vérifier que la ligne inscrite n'est pas vide
            if (!empty($row_content)) {
                if (!empty($key_line) && array_key_exists($key_line, $line)) {
                    $content[$line[$key_line]] = $line;
                } else {
                    $content[] = $line;
                }
            }
        }
        fclose($handle);

        return $content;
    }


    /**
     * @param string $file  Filepath or "php://output" stream
     * @param array $data
     * @param bool $first_line_title
     * @param string $delimiter
     * @param string $enclosure
     * @return bool
     */
    public static function generatecsv(
        string $file,
        array $data,
        bool $first_line_title = true,
        string $delimiter = CsvFile::DELIMITER,
        string $enclosure = CsvFile::ENCLOSURE,
    ): bool {
        $csv = new CsvFile($file);
        $csv->setOption($delimiter, $enclosure);

        if ('php://output' === $file && !headers_sent()) {
            $filename = str_replace('.php', '.csv', basename($_SERVER['PHP_SELF']));
            header('Content-Type: text/csv');
            header('Content-Disposition: attachment; filename="' . $filename . '"');
        }

        if ($first_line_title) {
            reset($data);
            $first_key = (string) key($data);

            $csv->writeTitle(array_keys($data[$first_key]));
        }

        foreach ($data as $fields) {
            $csv->write($fields);
        }

        $csv->close();

        return file_exists($file);
    }


    static public function content_to_database(
        \PDO $pdo,
        array $title,
        array $data,
        string $table,
        bool $create_table = false,
    ): bool {
        if(empty($title) || empty($data)) {
            throw new InvalidArgumentException("The arguments title or content are incorrect, require array no empty.");
        }

        if($create_table) {
            // Add a auto increment ID field
            // Your CSV file can have a "id" column or not, this field will be created on this table in all cases
            $sqlFields = ['`id` INT UNSIGNED NOT NULL AUTO_INCREMENT'];

            for($i = 0; $i < count($title); $i++) {
                $title[$i] = str_replace('-', '_', self::slugify($title[$i]));
                if ($title[$i] == 'id') {
                    continue; // Field already created
                }

                $sqlFields[] = "`{$title[$i]}` TEXT NULL DEFAULT NULL";
            }

            $sqlFields[] = 'PRIMARY KEY (`id`)';

            $sqlFiels = implode(', ', $sqlFields);
            $sql = "CREATE TABLE `{$table}` ( {$sqlFiels} ) ENGINE=InnoDB CHARSET=utf8;";
            $pdo->prepare($sql)->execute();
        }

        // Prepare INSERT
        $sqlFields = '`'.implode('`, `', $title).'`';
        $placeholder = substr(str_repeat('?,', count($title)),0,-1);

        $sql = "INSERT INTO `{$table}` ({$sqlFields}) VALUES ({$placeholder})";

        try {
            $pdo->beginTransaction();
            foreach($data as $content) {
                $pdo->prepare($sql)->execute($content);
            }

            return $pdo->commit();

        } catch (\PDOException $exception) {
            $pdo->rollBack();
            throw $exception;
        }
    }

    private static function csvContentExtractField(string $value, array $mode = self::EXTRACT_VALUE_DEFAULT_MODE): ?string
    {
        if (in_array(self::EXTRACT_VALUE_TRIM, $mode)) {
            $value = trim($value);
        }

        if (in_array(self::EXTRACT_VALUE_EMPTY_IS_NULL, $mode) && empty($value) && $value !== 0) {
            $value = null;
        } elseif (in_array(self::EXTRACT_VALUE_NULL_STRING, $mode) && in_array($value, ['null', 'NULL'])) {
            $value = null;
        }

        return $value;
    }

    /**
     * Converts a string into a slug. http://en.wikipedia.org/wiki/Slug_(web_publishing)#Slug
     * Source: https://gist.github.com/Narno/6540364 (Narno)
     */
    private static function slugify(string $string, string $separator = '-'): string
    {
        $string = preg_replace('/
		[\x09\x0A\x0D\x20-\x7E]              # ASCII
		| [\xC2-\xDF][\x80-\xBF]             # non-overlong 2-byte
		|  \xE0[\xA0-\xBF][\x80-\xBF]        # excluding overlongs
		| [\xE1-\xEC\xEE\xEF][\x80-\xBF]{2}  # straight 3-byte
		|  \xED[\x80-\x9F][\x80-\xBF]        # excluding surrogates
		|  \xF0[\x90-\xBF][\x80-\xBF]{2}     # planes 1-3
		| [\xF1-\xF3][\x80-\xBF]{3}          # planes 4-15
		|  \xF4[\x80-\x8F][\x80-\xBF]{2}     # plane 16
		/', '', $string);
        // @see https://github.com/cocur/slugify/blob/master/src/Cocur/Slugify/Slugify.php
        // transliterate
        $string = iconv('utf-8', 'us-ascii//TRANSLIT', $string);
        // replace non letter or digits by seperator
        $string = preg_replace('#[^\\pL\d]+#u', $separator, $string);
        // trim
        $string = trim($string, $separator);
        // lowercase
        $string = (defined('MB_CASE_LOWER')) ? mb_strtolower($string) : strtolower($string);
        // remove unwanted characters
        $string = preg_replace('#[^-\w]+#', '', $string);

        return $string;
    }

    private static function forceUtf8(string $text): string
    {
        if (!mb_detect_encoding($text, 'UTF-8', true)) {
            $text = mb_convert_encoding($text, 'UTF-8', mb_list_encodings());
        }

        return $text;
    }
}
