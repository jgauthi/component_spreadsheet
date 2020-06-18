<?php
/*******************************************************************************
 * @name: Export XML Excel
 * @note: Framework pour générer des fichiers excels en XML
 * @author: Jgauthi <https://github.com/jgauthi>, created the [21dec2015]
 * @version: 1.0
 * @Requirements:
    - PHP version >= 7.4+ (http://php.net)

 *******************************************************************************/

namespace Jgauthi\Component\Spreadsheet;

use DateTime;
use DateTimeInterface;
use Exception;
use InvalidArgumentException;

class ExcelXmlFile
{
    protected string $file;
    protected int $nb_col = 0;
    private int $chmod = 0664;

    /** @var false|resource */
    protected $stream;

    /**
     * File creation
     *
     * @param string $file File location: if it already exists, it will be deleted. Indicate 'php://output' to go through the stream server
     * @param int|null $chmod
     */
    public function __construct(string $file, ?int $chmod = null)
    {
        $this->file = $file;

        $this->stream = fopen($this->file, 'w');
        if (!$this->stream) {
            throw new InvalidArgumentException(sprintf('Impossible de créer le fichier "%s"', $file));
        }

        if ('php://output' === $this->file && !headers_sent()) {
            $filename = preg_replace("#(\.[^$]{2,5})$#i", '.xlsx.xml', basename($_SERVER['PHP_SELF']));

            header('Content-Disposition: attachment; filename=' . $filename);
            header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            header('Content-Transfer-Encoding: binary');
            header('Cache-Control: must-revalidate');
            header('Pragma: public');

            $this->chmod = 0;

        // If CHMOD is null, keep default value
        // else define value. 0 value mean chmod is disabled
        } elseif ($chmod !== null) {
            $this->chmod = $chmod;
        }
    }

    /**
     * Header of generated XML
     */
    public function header(string $authorName, ?DateTimeInterface $dateCreation = null, ?DateTimeInterface $dateUpdate = null): self
    {
        if(!$dateCreation) {
            $dateCreation = new DateTime;
        }
//        $dateCreation = $dateCreation->format('Y-m-d\TH:i:s\Z');

        if(!$dateUpdate) {
            $dateUpdate = new DateTime;
        }
//        $dateUpdate = $dateUpdate->format('Y-m-d\TH:i:s\Z');

        $header = <<<XML
<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:html="http://www.w3.org/TR/REC-html40">
 <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">
  <Author>{$authorName}</Author>
  <LastAuthor>{$authorName}</LastAuthor>
  <Created>{$dateCreation->format('Y-m-d\TH:i:s\Z')}</Created>
  <LastSaved>{$dateUpdate->format('Y-m-d\TH:i:s\Z')}</LastSaved>
  <Version>12.00</Version>
 </DocumentProperties>
 <ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">
  <WindowHeight>11505</WindowHeight>
  <WindowWidth>28455</WindowWidth>
  <WindowTopX>240</WindowTopX>
  <WindowTopY>120</WindowTopY>
  <ProtectStructure>False</ProtectStructure>
  <ProtectWindows>False</ProtectWindows>
 </ExcelWorkbook>
 <Styles>
  <Style ss:ID="Default" ss:Name="Normal">
   <Alignment ss:Vertical="Bottom"/>
   <Borders/>
   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"/>
   <Interior/>
   <NumberFormat/>
   <Protection/>
  </Style>
  <Style ss:ID="s65">
   <Alignment ss:Horizontal="Center" ss:Vertical="Top"/>
   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000" ss:Bold="1"/>
   <Interior ss:Color="#dddddd" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="s62">
   <NumberFormat ss:Format="@"/>
  </Style>
 </Styles>
 <Worksheet ss:Name="Export">
  <Names><NamedRange ss:Name="_FilterDatabase" ss:Hidden="1"/></Names>
  <Table>
   <Column ss:Index="3" ss:StyleID="s62"/>
XML;

        $this->write($header);
        return $this;
    }

    /**
     * XML footer, close the stream
     */
    public function footer(): bool
    {
        // Activate auto filter + line freeze if a title line has been added
        $filter_auto = $title_freeze = null;
        if (!empty($this->nb_col)) {
            $filter_auto = '<AutoFilter x:Range="R1C1:R1C' . $this->nb_col . '" xmlns="urn:schemas-microsoft-com:office:excel"></AutoFilter>';
            $title_freeze = '<FreezePanes/><FrozenNoSplit/><SplitHorizontal>1</SplitHorizontal><TopRowBottomPane>1</TopRowBottomPane><ActivePane>2</ActivePane>';
        }

        $footer = <<<XML
</Table>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <PageSetup>
    <Header x:Margin="0.3"/>
    <Footer x:Margin="0.3"/>
    <PageMargins x:Bottom="0.75" x:Left="0.7" x:Right="0.7" x:Top="0.75"/>
   </PageSetup>
   <Print>
    <ValidPrinterInfo/>
    <PaperSizeIndex>9</PaperSizeIndex>
    <HorizontalResolution>-3</HorizontalResolution>
    <VerticalResolution>0</VerticalResolution>
   </Print>
   <Selected/>
   {$title_freeze}
   <ProtectObjects>False</ProtectObjects>
   <ProtectScenarios>False</ProtectScenarios>
  </WorksheetOptions>
  {$filter_auto}
 </Worksheet>
</Workbook>
XML;

        $this->write($footer);
        if (!empty($this->chmod)) {
            chmod($this->file, $this->chmod);
        }

        return fclose($this->stream);
    }

    // Generation functions -----------------------------------------------------------

    /**
     * Convert a 2-dimensional array to XML
     * @throws Exception|InvalidArgumentException
     */
    public function array_to_xml(array $data, ?string $author = null): bool
    {
        if (empty($data)) {
            throw new InvalidArgumentException('Data is empty');
        }

        // Var
        reset($data);
        $first_key = key($data);
        if (empty($data[$first_key])) {
            return false;
        }

        // Table data
        $title = array_keys($data[$first_key]);
        if (empty($author)) {
            $author = 'Generated by ' . __CLASS__;
        }

        // Init XML
        $this
            ->header($author)
            ->writeTitle($title);

        foreach ($data as $line) {
            $this->lineOpen();
            foreach ($line as $col) {
                $this->writeCol($col);
            }

            $this->lineClose();
        }

        return $this->footer();
    }

    //-- Write function ---------------------------------------------------------------

    /**
     * Writes content to the current file
     */
    protected function write(string $content): self
    {
        fwrite($this->stream, $content);

        return $this;
    }

    /**
     * Open a row to allow adding columns
     */
    public function lineOpen(): self
    {
        return $this->write("\t<Row>\n");
    }

    /**
     * Close the current line
     */
    public function lineClose(): self
    {
        return $this->write("\t</Row>\n");
    }

    /**
     * Writes a column in the current line, needs to be called between line_open() and line_close()
     *
     * @param string $content
     * @param string $type String / Number
     */
    public function writeCol(string $content, string $type = 'String'): self
    {
        return $this->write(
            "\t\t".
            '<Cell><Data ss:Type="'. $type .'">'.
            $this->nl2br(htmlspecialchars($content, ENT_QUOTES, 'UTF-8')).
            '</Data></Cell>'."\n"
        );
    }

    /**
     * Write the title line
     */
    public function writeTitle(array $title): self
    {
        $this->write("\t<Row ss:StyleID=\"s65\">\n"); //$this->line_open();
        foreach ($title as $libelle) {
            $this->write(
                "\t\t".
                '<Cell><Data ss:Type="String">'.
                $this->nl2br(htmlspecialchars($libelle, ENT_QUOTES, 'UTF-8')).
                '</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>'."\n"
            );
        }

        // Keep the column number for the automatic filter
        $this->nb_col = count($title);

        return $this->lineClose();
    }

    /**
     * New line specific to XML Excel
     */
    private function nl2br(string $char): string
    {
        return str_replace(["\r\n", "\r", "\n"], '&#10;', trim($char));
    }
}
