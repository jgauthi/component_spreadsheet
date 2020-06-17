<?php
/*******************************************************************************
 * @name: Export Excel
 * @note: Framework to manage various aspects of the phpExcel class
 * @author: Jgauthi <https://github.com/jgauthi>, created the [19mai2012]
 * @version: 1.84
 * @Requirements:
    - Class phpExcel version >= 1.7.6 (http://phpexcel.codeplex.com)
    - PHP version >= 5.2.0 (http://fr2.php.net)
    - PHP extension: php_zip, php_xml, php_gd2 (if not compiled in)

 *******************************************************************************/
namespace Jgauthi\Component\Spreadsheet;

use PHPExcel;

class ExportExcel
{
    public $format_document = null;
    public $objEdition = null;
    protected $export_file = null;
    protected $tmp_dir = null;
    protected $sheet_selected = null;
    protected $dateformat = 'dd/mm/yyyy';

    public $cols = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z',
        'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ',
        'BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ', 'BK', 'BL', 'BM', 'BN', 'BO', 'BP', 'BQ', 'BR', 'BS', 'BT', 'BU', 'BV', 'BW', 'BX', 'BY', 'BZ',
        'CA', 'CB', 'CC', 'CD', 'CE', 'CF', 'CG', 'CH', 'CI', 'CJ', 'CK', 'CL', 'CM', 'CN', 'CO', 'CP', 'CQ', 'CR', 'CS', 'CT', 'CU', 'CV', 'CW', 'CX', 'CY', 'CZ', ];

    public $excel_info = [
        'creator' => 'Jgauthi',
        'subject' => 'Data_process_DO_NOT_DELETE_MENTION',
        'description' => null,
        'keywork' => 'office 2007',
        'category' => 'data csv',
    ];

    public $style = [
        'title' => [
            'font' => [
                'bold' => true,
            ],
            'alignment' => [
                'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_RIGHT,
            ],
            'borders' => [
                'top' => [
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                ],
            ],
            'fill' => [
                'type' => PHPExcel_Style_Fill::FILL_GRADIENT_LINEAR,
                'rotation' => 90,
                'startcolor' => [
                    'argb' => 'FFA0A0A0',
                ],
                'endcolor' => [
                    'argb' => 'FFFFFFFF',
                ],
            ],
        ],

        'line' => [
            'fill' => [
                'type' => PHPExcel_Style_Fill::FILL_GRADIENT_LINEAR,
                'rotation' => 90,
                'startcolor' => [
                    'argb' => 'FFD0D0D0',
                ],
                'endcolor' => [
                    'argb' => 'FFD0D0D0',
                ],
            ],
        ],
    ];

    public function __construct($type = null)
    {
        // Déclaration dossier temporaire
        $getenv_tmp_dir = getenv('TMP');

        if (defined('PATH_TMP_DIR')) {
            $this->tmp_dir = PATH_TMP_DIR;
        } elseif (defined('PATH_LOG_DIR')) {
            $this->tmp_dir = PATH_LOG_DIR;
        } elseif (function_exists('sys_get_temp_dir')) {
            $this->tmp_dir = sys_get_temp_dir();
        } elseif (!empty($getenv_tmp_dir)) {
            $this->tmp_dir = $getenv_tmp_dir;
        }

        if (empty($this->tmp_dir)) {
            $this->tmp_dir = __DIR__.'/../log/';
        }

        // Vérification dossier temporaire
        if (!is_writable($this->tmp_dir)) {
            return exit(!user_error("ECHEC: le répertoire temporaire '{$this->tmp_dir}' n'est pas disponible en écriture"));
        } elseif (!is_readable($this->tmp_dir)) {
            return exit(!user_error("ECHEC: le répertoire temporaire '{$this->tmp_dir}' n'est pas disponible en lecture"));
        }

        // OK
        if (empty($getenv_tmp_dir)) {
            putenv('TMP='.$this->tmp_dir);
        }

        // Définir le type du document
        if (in_array($type, [2007, 'excel2007', 'xlsx'])) {
            $this->format_document = ['version' => 'Excel2007', 'ext' => 'xlsx', 'mime' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'];
        } else {
            $this->format_document = ['version' => 'Excel5', 'ext' => 'xls', 'mime' => 'application/vnd.ms-excel'];
        }
    }

    //-- Fonctions d'import / export ---------------------------------------------------------------------------------------------------
    public function array_to_excel($filename, $data, $cols_width = 'auto')
    {
        // Initialisation class
        $this->objEdition = new PHPExcel();
        $this->objEdition->getProperties()
            ->setCreator($this->excel_info['creator'])
            ->setLastModifiedBy($this->excel_info['creator'])
            ->setTitle($filename)
            ->setSubject($this->excel_info['subject'])
            ->setDescription($this->excel_info['description'])
            ->setKeywords($this->excel_info['keywork'])
            ->setCategory($this->excel_info['category']);

        for ($i = 0; $i < count($data); ++$i) {
            // Créer un nouvel onglet
            if ($i > 0) {
                $this->objEdition->createSheet();
            }
            $this->objEdition->setActiveSheetIndex($i);
            $this->objEdition->getActiveSheet()->setTitle($data[$i]['name']);
            //$this->objEdition->getActiveSheet()->getProtection()->setSheet(true);
            $line = 1;

            // Titre
            if (!empty($data[$i]['title'])) {
                for ($j = 0; $j < count($data[$i]['title']); ++$j) {
                    $this->objEdition->getActiveSheet()->setCellValue($this->cols[$j].$line, $data[$i]['title'][$j]);
                }

                // Style titre
                $this->objEdition->getActiveSheet()->getStyle('A1:'.$this->cols[(count($data[$i]['title']) - 1)].'1')->applyFromArray($this->style['title']);

                ++$line;
            }

            // Contenu
            if (!empty($data[$i]['content'])) {
                foreach ($data[$i]['content'] as $content_line) {
                    //-- LoopCallback function (issue de vyset, à implémenter ?)
                    /*if(!empty($this->loopCallback) && is_callable($this->loopCallback['callback'])) {
                        call_user_func($this->loopCallback['callback'], $contentKey, $data[$i]['content'],$this->loopCallback['params']);
                    }*/

                    // Insérer les colonnes de la ligne
                    if (empty($content_line)) {
                        continue;
                    }

                    // Imposer les keys du tableau en numérique
                    elseif (!isset($content_line[0])) {
                        $content_line = array_values($content_line);
                    }

                    // Insérer les colonnes de la ligne
                    foreach ($content_line as $index => $content) {
                        // Format de colonne personnalisée
                        if (!empty($data[$i]['type_colonne'][$index])) {
                            switch ($data[$i]['type_colonne'][$index]) {
                                case 'int': $this->insert_datacell_integer($this->cols[$index].$line, $content); break;
                                case 'float': $this->insert_datacell_float($this->cols[$index].$line, $content); break;
                                case 'pourcent': $this->insert_datacell_pourcent($this->cols[$index].$line, $content); break;
                                case 'date': $this->insert_datacell_date($this->cols[$index].$line, $content); break;
                                case 'text': $this->insert_datacell_text($this->cols[$index].$line, $content); break;

                                // Colonne non reconnu
                                default:
                                    return die(!user_error("Colonne '{$data[$i]['type_colonne'][$index]}' non reconnu"));
                                    break;
                            }
                        } else {
                            if (!mb_detect_encoding($content, 'UTF-8', true)) {
                                $content = utf8_encode($content);
                            }

                            $this->objEdition->getActiveSheet()->setCellValue($this->cols[$index].$line, $content);
                        }
                    }

                    // Protection (trop fonctionnel)
                    //$this->objEdition->getActiveSheet()->protectCells("A{$line}:". $this->cols[($index - 1)] .$line, 'k3nnEkEoW');

                    // Fond gris zèbre
                    if (0 !== $line % 2) {
                        $this->objEdition->getActiveSheet()->getStyle("A{$line}:{$this->cols[$index]}{$line}")->applyFromArray($this->style['line']);
                    }

                    ++$line;
                }
            }

            // Style titre-colonne
            if (!empty($this->style['column_title'])) {
                $this->objEdition->getActiveSheet()->getStyle('A1:'.'A'.($line - 1))->applyFromArray($this->style['column_title']);
            }

            // Largueur des colonnes
            if (!empty($cols_width)) {
                if ('auto' === $cols_width) {
                    for ($k = 0; $k < $index; ++$k) {
                        $this->objEdition->getActiveSheet()->getColumnDimension($this->cols[$k])->setAutoSize(true);
                    }
                } elseif (is_array($cols_width)) {
                    foreach ($cols_width as $id_col => $value_cols) {
                        $this->objEdition->getActiveSheet()->getColumnDimension($this->cols[$id_col])->setWidth($value_cols);
                    }
                }
            }
        }

        // Enregistrer le fichier
        if (!is_null($filename)) {
            return $this->save_file($filename);
        }
    }

    public function array_to_excel_vygon($filename, $data, $cols_width = 'auto') // Sera renommé array_to_excel à la 2.0
    {
        user_error('La fonction '.__FUNCTION__.'($filename) est deprecated, utiliser de préférence: array_to_excel($filename, $data)');
        $this->array_to_excel($filename, $data, $cols_width);
    }

    public function excel_to_array($file, $first_line_title = false)
    {
        if (!$this->excel_check_file($file)) {
            return false;
        }

        $objPHPExcel = PHPExcel_IOFactory::load($file);
        $excel = [];
        $i = 0;

        foreach ($objPHPExcel->getWorksheetIterator() as $worksheet) {
            $excel[$i] = [
                'name' => $worksheet->getTitle(),
                'letter_max' => $worksheet->getHighestColumn(),
                'row' => $worksheet->getHighestRow(),
                'col' => ord($worksheet->getHighestColumn()) - 64,
                'type_colonne' => ['text'],
                'title' => [],
                'content' => [],
            ];
            $highestColumnIndex = PHPExcel_Cell::columnIndexFromString($excel[$i]['letter_max']);

            for ($row = 1; $row <= $excel[$i]['row']; ++$row) {
                $length = 0;
                $ligne = [];

                // Récupérer la ligne complète
                for ($col = 0; $col < $highestColumnIndex; ++$col) {
                    $val = trim($worksheet->getCellByColumnAndRow($col, $row)->getFormattedValue());
                    //$val = $worksheet->getCellByColumnAndRow($col, $row)->getFormattedValue();
                    $length += strlen($val);

                    $ligne[] = $val;
                }

                // Stocker les données
                if ($length > 0) {
                    if (1 === $row && $first_line_title) {
                        $excel[$i]['title'] = $ligne;
                    } else {
                        $excel[$i]['content'][] = $ligne;
                    }
                } else {
                    $excel[$i]['row']--;
                }
            }

            ++$i;
        }

        return $excel;
    }

    // Récupère les données d'un onglet d'un fichier excel, retourne un tableau sous une forme équivalente d'un mysql_fetch_assoc
    public function excel_fetch_array($onglet, $first_line_title = true, $colonne_index = null)
    {
        if (is_null($this->objEdition)) {
            return die(!user_error('Fonction '.__FUNCTION__.": ne peut requêter l'onglet '$onglet', aucun fichier ouvert"));
        }

        // Ouverture de l'onglet
        $worksheet = $this->objEdition->getSheetByName($onglet);
        if (is_null($worksheet)) {
            return die(!user_error('Fonction '.__FUNCTION__.": ne peut requêter l'onglet '$onglet', celui-ci semble ne pas exister."));
        }

        $excel = [
            'letter_max' => $worksheet->getHighestColumn(),
            'row' => $worksheet->getHighestRow(),
            'col' => ord($worksheet->getHighestColumn()) - 64,
        ];
        $highestColumnIndex = PHPExcel_Cell::columnIndexFromString($excel['letter_max']);

        // Récupération des données
        $data = [];
        $title = null;

        for ($row = 1; $row <= $excel['row']; ++$row) {
            $length = 0;
            $ligne = [];

            // Gestion du titre
            if (1 === $row) {
                for ($col = 0; $col < $highestColumnIndex; ++$col) {
                    $val = trim($worksheet->getCellByColumnAndRow($col, $row)->getFormattedValue()); // getCalculatedValue
                    //$val = $worksheet->getCellByColumnAndRow($col, $row)->getFormattedValue(); // getCalculatedValue
                    $length += strlen($val);

                    $ligne[] = $val;
                }

                // Déterminer la ligne de titre, ainsi que la colonne qui pourrait servir d'index
                if ($first_line_title) {
                    $title = $ligne;
                } else {
                    $data[] = $ligne;
                }

                // Désactiver
                if (!is_array($title) || !in_array($colonne_index, $title)) {
                    $colonne_index = null;
                }

                continue;
            }

            // Récupérer la ligne complète
            for ($col = 0; $col < $highestColumnIndex; ++$col) {
                $val = trim($worksheet->getCellByColumnAndRow($col, $row)->getFormattedValue());
                //$val = $worksheet->getCellByColumnAndRow($col, $row)->getFormattedValue();
                $length += strlen($val);

                if (!is_null($title)) {
                    $ligne[$title[$col]] = $val;
                } else {
                    $ligne[] = $val;
                }
            }

            // Stocker les données
            if ($length > 0) {
                if (!is_null($colonne_index)) {
                    $data[$ligne[$colonne_index]] = $ligne;
                } else {
                    $data[] = $ligne;
                }
            } else {
                $excel['row']--;
            }
        }

        // Envoyer les données
        return (!empty($data)) ? $data : null;
    }

    // Récupére la liste des données sur une colonne (A, B...)
    public function excel_get_value_col($sheet, $col_selected, $first_line = 1)
    {
        if (is_null($this->objEdition)) {
            return die(!user_error('Fonction '.__FUNCTION__.": ne peut requêter l'onglet '$onglet', aucun fichier ouvert"));
        } elseif (!is_numeric($sheet)) {
            return die(!user_error('Fonction '.__FUNCTION__.": ne peut requêter l'onglet '$onglet', ".
                "indiquer l'index numérique de l'onglet (vous pouvez utiliser get_index_sheet($sheet_name pour récupérer son index)"));
        }

        // Ouverture de l'onglet
        $worksheet = $this->objEdition->getSheet($sheet);

        // Récupération des données
        $col_selected = array_search($col_selected, $this->cols);
        $data = [];

        for ($row = intval($first_line); $row <= $worksheet->getHighestRow(); ++$row) {
            $val = trim($worksheet->getCellByColumnAndRow($col_selected, $row)->getCalculatedValue());
            //$val = $worksheet->getCellByColumnAndRow($col_selected, $row)->getCalculatedValue();

            if ('' !== $val) {
                $data[$row] = $val;
            }
        }

        // Envoyer les données
        return (!empty($data)) ? $data : null;
    }

    //-- Gestion des onglets ------------------------------------------------------------------------------------------

    // Sélectionne un onglet si nécessaire (nécessaire avant d'insérer des données)
    public function select_sheet($sheet_id)
    {
        if ($sheet_id === $this->sheet_selected) {
            return true;
        }

        $this->objEdition->setActiveSheetIndex($sheet_id);
    }

    // Récupère l'index d'un onglet à partir de son nom
    public function get_index_sheet($sheet_name)
    {
        $index = null;
        if (($worksheet = $this->objEdition->getSheetByName($sheet_name)) instanceof PHPExcel_Worksheet) {
            $index = $worksheet->getParent()->getIndex($worksheet);
        }

        return  $index;
    }

    // Suppression d'un onglet à partir de son index
    public function excel_delete_sheet($sheet_id)
    {
        if (is_null($this->objEdition)) {
            return die(!user_error('Fonction '.__FUNCTION__.": ne peut requêter l'onglet num{$sheet_id}, aucun fichier ouvert"));
        }

        return  $this->objEdition->removeSheetByIndex($sheet_id);
    }

    //-- Fonctions d'édition de fichiers ------------------------------------------------------------------------------

    // Vérifier que le fichier est accessible (lisible en écriture et non verrouillé)
    public function excel_check_file($filename)
    {
        // Check que le fichier existe et est lisible
        if (!is_readable($filename)) {
            return !user_error("Le fichier '$filename' n'existe pas ou n'est pas lisible en lecture");
        }

        // Verifier que le fichier n'est pas déjà verrouillé par une autre application
        elseif (is_writable($filename) && $fp = @fopen($filename, 'r+')) {
            // Le fichier est verrouillé
            if (!flock($fp, LOCK_EX | LOCK_NB)) {
                fclose($fp);

                return !user_error("Le fichier '$filename' est actuellement verrouillé car utilisé par une autre application");
            }

            flock($fp, LOCK_UN);
            fclose($fp);
        }

        // Vérifier le type du fichier
        $check = PHPExcel_IOFactory::identify($filename);
        if (!in_array($check, ['Excel5', 'Excel2007'])) {
            return !user_error("Le format de document '$check' n'est pas supporté pour le fichier '$filename'");
        }

        return true;
    }

    public function excel_import_file($filename)
    {
        if ($this->excel_check_file($filename)) {
            $this->objEdition = PHPExcel_IOFactory::load($filename);

            return true;
        }

        return false;
    }

    public function save_file($filename)
    {
        if (empty($this->objEdition)) {
            return false;
        }

        $this->file_export = $filename;

        // Supprimer l'ancien fichier
        $new_file = true;
        if (file_exists($this->file_export) && !@unlink($this->file_export)) {
            if (!is_writable($this->file_export)) {
                return !user_error("Impossible de sauvegarder les modifications, le fichier '{$this->file_export}' est n'est pas accessible en écriture");
            }

            return !user_error("Impossible de sauvegarder les modifications, le fichier '{$this->file_export}' est verrouillé par une autre application");

            $new_file = false;
        }

        $this->select_sheet(0);
        $objWriter = PHPExcel_IOFactory::createWriter($this->objEdition, $this->format_document['version']);
        $objWriter->save($this->file_export);

        // Erreur lors de la génération :(
        if (file_exists($this->file_export)) {
            if ($new_file) {
                @chmod(0644, $this->file_export);
            }

            return true;
        }

        return exit(!user_error("ECHEC lors de la création du fichier: '{$this->file_export}'"));
    }

    // [DEPRECATED] Alias de save_file(), maintenue pour rester retro-compatible
    public function excel_save_imported_file($filename)
    {
        user_error('La fonction '.__FUNCTION__.'($filename) est deprecated, utiliser de préférence: save_file($filename)');

        return $this->save_file($filename);
    }

    public function download_file($file_export)
    {
        if (headers_sent()) {
            return exit(!user_error('Erreur détecté durant l\'execution du script, fin de parcours'));
        }

        // Créer le fichier d'export si il n'existe pas
        if (is_null($this->export_file)) {
            // Infos de base pour le téléchargement
            $filename = basename($file_export);

            // Stocker temporairement l'export XLS
            $file_export = $this->tmp_dir.'/export_xls_'.date('YmdHis').'.xls.tmp';
            $this->save_file($file_export);
            $taille = filesize($file_export);

            // Prévoir la suppression du fichier temporaire à la fin du script
            register_shutdown_function(create_function('', "@unlink('$file_export');"));
        // $objWriter->save('php://output');
        } else {
            $filename = basename($file_export);
            $taille = filesize($file_export);
        }

        // Lancer le téléchargement
        header("Content-Type: {$this->format_document['mime']}; name=\"$filename\""); // application/force-download
        header('Content-Transfer-Encoding: binary');
        header("Content-Length: $taille");
        header("Content-Disposition: attachment; filename=\"$filename\"");
        header('Expires: 0');
        header('Cache-Control: no-cache, must-revalidate');
        header('Pragma: no-cache');
        readfile($file_export);

        return exit();
    }

    // [DEPRECATED] Alias de download_file(), maintenue pour rester retro-compatible
    public function excel_download_imported_file($file_export)
    {
        user_error('La fonction '.__FUNCTION__.'($filename) est deprecated, utiliser de préférence: download_file($filename)');

        return $this->download_file($file_export);
    }

    // Fermer et effacer les données en mémoire en cours
    public function excel_close()
    {
        $this->objEdition->disconnectWorksheets();
        $this->objEdition = null;
    }

    /**
     * Suppression des dossiers temporaires crées par PHP Excel.
     *
     * @param string $filtre Filtre date interprété par strtotime (défaut: first day of -1 month)
     *
     * @return (array|void)
     */
    public function delete_old_tmpdir($filtre = 'first day of -1 month')
    {
        $time = strtotime($filtre);
        if (empty($time) || !function_exists('glob')) {
            return !user_error(sprintf('In %s: function glob no exists.', __CLASS__.'->'.__FUNCTION__));
        } elseif (empty($this->tmp_dir)) {
            return !user_error(sprintf('In %s: empty tmpdir.', __CLASS__.'->'.__FUNCTION__));
        } elseif (!function_exists('delete_directory')) {
            return !user_error(sprintf('In %s: function delete_directory require, no exist.', __CLASS__.'->'.__FUNCTION__));
        }

        // Récupérer la liste des dossiers temporaires
        $dir_liste = @glob("{$this->tmp_dir}/tmp_*", GLOB_ONLYDIR);
        if (empty($dir_liste)) {
            return false;
        }

        $nb = ['delete' => 0, 'failed' => 0];
        foreach ($dir_liste as $dir) {
            $date_update = filemtime($dir);
            if ($date_update < $time) {
                if (@delete_directory($dir, true)) {
                    ++$nb['delete'];
                } else {
                    $nb['failed']++;
                }
            }
        }

        return $nb;
    }

    //-- Fonctions gestion ligne/colonnes dans l'onglet en cours --------------------------------------------------------
    public function format_colonne($col)
    {
        // Format texte supporté (pour l'instant)
        $type = PHPExcel_Style_NumberFormat::FORMAT_TEXT;

        $this->objEdition->getActiveSheet()->getStyle($col)->getNumberFormat()->setFormatCode($type);
    }

    // Format date (conversion au format excel, ex: d/m/Y => dd/mm/yy), pour les fonctions insert_date & insert_celldara_date
    public function format_date($dateformat)
    {
        $this->dateformat = str_replace(['d', 'm', 'y', 'Y'], ['dd', 'mm', 'yy', 'yyyy'], $dateformat);
    }

    // Permet de fixer une ligne et/ou une colonne pour les conserver lors du défilement horizontal/vertical de la feuille excel
    public function fixer_contenu($ligne = null, $col = null)
    {
        // Incrémenter lu numero de ligne et de colonne de +1 (bug ?)
        $ligne = intval($ligne) + 1;
        $col = ((!is_null($col)) ? $this->cols[array_search($col, $this->cols) + 1] : 'A');

        $this->objEdition->getActiveSheet()->freezePane($col.$ligne);
    }

    //-- Fonctions pour remplir et formater les champs dans un fichier en cours d'édition -------------------------------

    // [DEPRECATED] Sera supprimé/refondu à la 2.0, utilisé de préférence insert_datacell_*
    public function insert_pourcent($ActiveSheet, $data)
    {
        if (!is_array($data) || empty($data) || is_null($this->objEdition)) {
            return false;
        }

        $this->select_sheet($ActiveSheet);
        foreach ($data as $colonne => $pourcent) {
            $this->insert_datacell_pourcent($colonne, $pourcent);
        }
    }

    // [DEPRECATED] Sera supprimé/refondu à la 2.0, utilisé de préférence insert_datacell_*
    public function insert_text($ActiveSheet, $data)
    {
        if (!is_array($data) || empty($data) || is_null($this->objEdition)) {
            return false;
        }

        $this->select_sheet($ActiveSheet);
        foreach ($data as $colonne => $text) {
            $this->insert_datacell_text($colonne, $text);
        }
    }

    // [DEPRECATED] Sera supprimé/refondu à la 2.0, utilisé de préférence insert_datacell_*
    public function insert_date($ActiveSheet, $data, $dateformat = null)
    {
        if (!is_array($data) || empty($data) || is_null($this->objEdition)) {
            return false;
        }

        // Format date (conversion au format excel, ex: d/m/Y => dd/mm/yy)
        if (!is_null($dateformat)) {
            $dateformat = $this->format_date($dateformat);
        }

        // Parser les dates et les inclure dans le excel
        $this->select_sheet($ActiveSheet);
        foreach ($data as $colonne => $date) {
            $this->insert_datacell_date($colonne, $date);
        }
    }

    // [DEPRECATED] Sera supprimé/refondu à la 2.0, utilisé de préférence insert_datacell_*
    public function insert_integer($ActiveSheet, $data, $separateur_millier = false)
    {
        if (!is_array($data) || empty($data) || is_null($this->objEdition)) {
            return false;
        }

        $this->select_sheet($ActiveSheet);
        foreach ($data as $colonne => $integer) {
            $this->insert_datacell_integer($colonne, $integer, $separateur_millier);
        }
    }

    // [DEPRECATED] Sera supprimé/refondu à la 2.0, utilisé de préférence insert_datacell_*
    public function insert_float($ActiveSheet, $data, $separateur_millier = false)
    {
        if (!is_array($data) || empty($data) || is_null($this->objEdition)) {
            return false;
        }

        $this->select_sheet($ActiveSheet);
        foreach ($data as $colonne => $float) {
            $this->insert_datacell_float($colonne, $float, $separateur_millier);
        }
    }

    // [DEPRECATED] Sera supprimé/refondu à la 2.0, utilisé de préférence insert_datacell_*
    public function insert_comment($ActiveSheet, $data, $author = null)
    {
        // Les commentaires sont seulement supportés sur la version 2007 de phpExcel
        if (!is_array($data) || empty($data) || is_null($this->objEdition)) {
            return false;
        }

        // Indiquer pour le système un nom d'auteur
        $excel_system_author_comment = ((!is_null($author)) ? $author : $this->excel_info['creator']);

        $this->select_sheet($ActiveSheet);
        foreach ($data as $colonne => $text) {
            $this->insert_datacell_comment($colonne, $text, $author);
        }
    }

    //-- Insertions de données typés  -----------------------------------------------------------------------------------------------------
    public function insert_datacell_integer($colonne, $integer, $separateur_millier = false)
    {
        if ($separateur_millier) {
            $type = '#,##0';
        } else {
            $type = PHPExcel_Style_NumberFormat::FORMAT_NUMBER;
        }

        $this->objEdition->getActiveSheet()->setCellValue($colonne, $integer)
            ->getStyle($colonne)->getNumberFormat()->setFormatCode($type);

        return $this;
    }

    public function insert_datacell_float($colonne, $float, $separateur_millier = false)
    {
        if ($separateur_millier) {
            $type = '#,##0.00';
        } else {
            $type = PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00;
        }

        $this->objEdition->getActiveSheet()->setCellValue($colonne, $float)
            ->getStyle($colonne)->getNumberFormat()->setFormatCode($type);

        return $this;
    }

    public function insert_datacell_pourcent($colonne, $pourcent)
    {
        if (!is_null($pourcent) && is_numeric($pourcent)) {
            $pourcent = $pourcent / 100;
        }

        $this->objEdition->getActiveSheet()->setCellValue($colonne, $pourcent)
            ->getStyle($colonne)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00);

        return $this;
    }

    public function insert_datacell_date($colonne, $date)
    {
        if (empty($date)) {
            return $this;
        } elseif (preg_match('#^([0-9]{4})(-|/)([0-9]{2})(-|/)([0-9]{2})$#', $date, $row)) {
            list(, $year, , $month, , $day) = $row;
        } elseif (preg_match('#^([0-9]{2})(-|/)([0-9]{2})(-|/)([0-9]{4})$#', $date, $row)) {
            list(, $day, , $month, , $year) = $row;
        } else {
            return exit(!user_error("Format de date incompatible: '$date', format souhaité: YYYY/MM/DD"));
        }

        // static long FormattedPHPToExcel( long $year, long $month, long $day, [long $hours = 0], [long $minutes = 0], [long $seconds = 0])
        $date = PHPExcel_Shared_Date::FormattedPHPToExcel($year, $month, $day);

        $this->objEdition->getActiveSheet()->setCellValue($colonne, $date)
            ->getStyle($colonne)->getNumberFormat()->setFormatCode($this->dateformat);

        return $this;
    }

    public function insert_datacell_text($colonne, $text)
    {
        if (!mb_detect_encoding($text, 'UTF-8', true)) {
            $text = utf8_encode($text);
        }

        $this->objEdition->getActiveSheet()->setCellValueExplicit($colonne, $text, PHPExcel_Cell_DataType::TYPE_STRING)
            ->getStyle($colonne)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_TEXT);

        return $this;
    }

    public function insert_datacell_comment($colonne, $text, $author = null)
    {
        static $excel_system_author_comment = null;

        // Définie l'auteur système (pour MS Excel) du commentaire
        if (is_null($excel_system_author_comment)) {
            $excel_system_author_comment = ((!is_null($author)) ? $author : $this->excel_info['creator']);
        }

        // Indiquer le nom de l'auteur dans le commentaire
        if (!is_null($author)) {
            $this->objEdition->getActiveSheet()->getComment($colonne)->getText()
                ->createTextRun("$author:\r\n")->getFont()->setBold(true);
        }

        // Ajouter le texte
        if (!mb_detect_encoding($text, 'UTF-8', true)) {
            $text = utf8_encode($text);
        }

        $this->objEdition->getActiveSheet()->getComment($colonne)
            ->setAuthor($excel_system_author_comment)
            ->getText()->createTextRun($text);

        return $this;
    }
}
