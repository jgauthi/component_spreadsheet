<?php
/*******************************************************************************
  * @name: CsvFile
  * @note: Framework for generating CSV files
  * @author: Jgauthi <https://github.com/jgauthi>, created the [23oct2018]
  * @version: 1.0
  * @Requirements:
        - PHP version >= 7.4+ (http://php.net)

*******************************************************************************/

namespace Jgauthi\Component\Spreadsheet;

use InvalidArgumentException;

class CsvFile
{
    public const DELIMITER = ';';
    public const ENCLOSURE = '"';

    protected string $file;
    protected array $options = [];
    private int $chmod = 0664;

    /** @var false|resource $stream */
    protected $stream;

    /**
     * File creation
     *
     * @param string $file File location: if it already exists, it will be deleted. Indicate 'php://output' to go through the stream server
     * @param int|null $chmod Set chmod file, you can disabled this feature with the value: 0
     */
    public function __construct(string $file, ?int $chmod = null)
    {
        $this->file = $file;

        $this->stream = fopen($this->file, 'w');
        if (!$this->stream) {
            throw new InvalidArgumentException(sprintf('Unable to create file "%s"', $file));
        }

        if ('php://output' === $this->file && !headers_sent()) {
            $filename = preg_replace("#(\.[^$]{2,5})$#i", '.csv', basename($_SERVER['PHP_SELF']));

            header('Content-Disposition: attachment; filename=' . $filename);
            header('Content-Type: text/csv');
            header('Content-Transfer-Encoding: UTF-8');
            header('Cache-Control: must-revalidate');
            header('Pragma: public');

            $this->chmod = 0;

        // If CHMOD is null, keep default value
        // else define value. 0 value mean chmod is disabled
        } elseif ($chmod !== null) {
            $this->chmod = $chmod;
        }

        // CSV Option by default
        $this->setOption();
    }

    public function setOption(string $delimiter = self::DELIMITER, string $enclosure = self::ENCLOSURE, bool $utf8_bom = false): self
    {
        $this->options = [
            'delimiter' => $delimiter,
            'enclosure' => $enclosure,
            'utf8_bom' => $utf8_bom,
        ];

        return $this;
    }

    public function writeTitle(array $titles = []): bool
    {
        if ($this->options['utf8_bom']) {
            fprintf($this->stream, chr(0xef) . chr(0xbb) . chr(0xbf));
        }

        return $this->write($titles);
    }

    public function write(array $fields): bool
    {
        return fputcsv(
            $this->stream,
            $fields,
            $this->options['delimiter'],
            $this->options['enclosure']
        );
    }

    // Close the stream
    public function close(): bool
    {
        if (!empty($this->chmod)) {
            chmod($this->file, $this->chmod);
        }

        return fclose($this->stream);
    }
}
