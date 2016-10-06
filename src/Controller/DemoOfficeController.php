<?php

namespace Fi\DemoBundle\Controller;

use Symfony\Bundle\FrameworkBundle\Controller\Controller;
use Symfony\Component\HttpFoundation\Response;
use PhpOffice\PhpWord\PhpWord;
use Symfony\Component\HttpKernel\Kernel;

class DemoOfficeController extends Controller
{
    public function PHPwordAction()
    {
        $doc = new PHPWord();
        //Aprire un file esistente
        $filepath = $this->get('kernel')->getRootDir().'/tmp/doc/attestato.docx';
        //File di destinazione
        $filepathnew = $this->get('kernel')->getRootDir().'/tmp/doc/attestato_new.docx';

        //Carica il template
        $document = $doc->loadTemplate($filepath);
        //Sostituisce [MATRICOLA] con il valore 59495
        $document->setValue('MATRICOLA', '59495');
        //Salva il documento sul nuovo file
        $document->save($filepathnew);

        return $this->render('DemoBundle:Demo:output.html.twig');
    }

    public function excelreadAction()
    {
        set_time_limit(960);
        ini_set('memory_limit', '2048M');

        $cacheMethod = \PHPExcel_CachedObjectStorageFactory::cache_to_phpTemp;
        $cacheSettings = array('memoryCacheSize' => '8MB');
        \PHPExcel_Settings::setCacheStorageMethod($cacheMethod, $cacheSettings);

        //Aprire un file esistente
        $filepath = $this->get('kernel')->getRootDir().'/tmp/rec.xls';
        $objPHPExcel = \PHPExcel_IOFactory::load($filepath);
        $sheet = $objPHPExcel->getActiveSheet();
        $totalRows = $sheet->getHighestRow();

        $matricolacol = 0;
        $cognomecol = 1;
        $nomecol = 2;
        $datanascitacol = 3;
        $datiletti = '';
        for ($row = 2; $row < $totalRows + 1; ++$row) {
            //Leggere il valore in una cella
            $matricola = $sheet->getCellByColumnAndRow($matricolacol, $row)->getValue();
            $cognome = $sheet->getCellByColumnAndRow($cognomecol, $row)->getValue();
            $nome = $sheet->getCellByColumnAndRow($nomecol, $row)->getValue();
            $dataf = $sheet->getCellByColumnAndRow($datanascitacol, $row)->getValue();
            $datanascita = \PHPExcel_Style_NumberFormat::toFormattedString($dataf, 'DD/MM/YYYY');

            //Leggere il valore del risultato di una formula
            /* $formula = $sheet->getCellByColumnAndRow($matricolacol, $row)->getCalculatedValue(); */

            $datiletti = $matricola.':'.$cognome.':'.$nome.':'.$datanascita;
        }

        //Read more: http://bayu.freelancer.web.id/2010/07/16/phpexcel-advanced-read-write-excel-made-simple/#ixzz2bGzPoFGk
        //Under Creative Commons License: Attribution

        return $this->render('DemoBundle:Demo:output.html.twig', array('extrainfo' => $datiletti));
    }

    public function excelreadarrayAction()
    {
        set_time_limit(960);
        ini_set('memory_limit', '2048M');

        $cacheMethod = \PHPExcel_CachedObjectStorageFactory::cache_to_phpTemp;
        $cacheSettings = array('memoryCacheSize' => '8MB');
        \PHPExcel_Settings::setCacheStorageMethod($cacheMethod, $cacheSettings);

        //Aprire un file esistente
        $filepath = $this->get('kernel')->getRootDir().'/tmp/rec.xls';
        $objPHPExcel = \PHPExcel_IOFactory::load($filepath);
        $sheet = $objPHPExcel->getActiveSheet();

        $values = $sheet->toArray();

        return $this->render('FiPhpExcelBundle:Default:read.html.twig', array('extrainfo' => $values));
    }

    public function excelwriteAction()
    {
        set_time_limit(960);
        ini_set('memory_limit', '2048M');

        $cacheMethod = \PHPExcel_CachedObjectStorageFactory::cache_to_phpTemp;
        $cacheSettings = array('memoryCacheSize' => '8MB');
        \PHPExcel_Settings::setCacheStorageMethod($cacheMethod, $cacheSettings);

        //Creare un nuovo file
        $objPHPExcel = new \PHPExcel();

        $objPHPExcel->setActiveSheetIndex(0);

        // Set properties
        $objPHPExcel->getProperties()->setCreator('Comune di Firenze');
        $objPHPExcel->getProperties()->setLastModifiedBy('Comune di Firenze');
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->setTitle('TitoloFoglio');

        // Si imposta il font
        //Times new romans
        $sheet->getDefaultStyle()->getFont()->setName('Verdana');
        $sheet->getDefaultStyle()->getFont()->setSize(12);

        //Si imposta la larghezza delle colonne
        $sheet->getColumnDimension('A')->setWidth(10);
        $sheet->getColumnDimension('B')->setWidth(25);
        $sheet->getColumnDimension('C')->setWidth(50);

        //Si imposta il colore dello sfondo delle celle
        //Colore header
        $style_header = array(
            'fill' => array(
                'type' => \PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => 'FFFF99'),
            ),
            'font' => array(
                'bold' => true,
                'color' => array('rgb' => 'FF0000'),
            ),
        );

        $sheet->getStyle('A1:C1')->applyFromArray($style_header);

        $col = 0;
        //Si scrive nelle celle
        $sheet->setCellValueByColumnAndRow($col, 1, 'NOME');
        $col = $col + 1;
        $sheet->setCellValueByColumnAndRow($col, 1, 'DATA DI NASCITA');
        $col = $col + 1;
        $sheet->setCellValueByColumnAndRow($col, 1, 'IMPORTO');
        $col = $col + 1;
        $sheet->setCellValueByColumnAndRow($col, 4, 'SOMMA');
        $col = $col + 1;

        //Ultima riga con valori
        //$sheet->getHighestRow()

        $col = 0;
        $row = 2;
        $sheet->setCellValueByColumnAndRow($col, $row, 'ANDREA MANZI');
        $col = $col + 1;

        $sheet->setCellValueByColumnAndRow($col, $row, '07/01/1980');
        \PHPExcel_Cell::setValueBinder(new \PHPExcel_Cell_DefaultValueBinder());
        $sheet->getStyle(\PHPExcel_Cell::stringFromColumnIndex($col).$row)->getNumberFormat()->setFormatCode('dd/mm/yyyy');

        $col = $col + 1;
        $sheet->setCellValueByColumnAndRow($col, $row, '2858.23');
        $sheet->getStyleByColumnAndRow($col, $row)->getNumberFormat()->setFormatCode('€ #,##0.00');

        $col = 0;
        $row = 3;
        $sheet->setCellValueByColumnAndRow($col, $row, 'ERIKA FENAROLI');
        $col = $col + 1;

        $sheet->setCellValueByColumnAndRow($col, $row, '15/01/1984');
        \PHPExcel_Cell::setValueBinder(new \PHPExcel_Cell_DefaultValueBinder());
        $sheet->getStyle(\PHPExcel_Cell::stringFromColumnIndex($col).$row)->getNumberFormat()->setFormatCode('dd/mm/yyyy');

        $col = $col + 1;
        $sheet->setCellValueByColumnAndRow($col, $row, '2945.89');
        $sheet->getStyleByColumnAndRow($col, $row)->getNumberFormat()->setFormatCode('€ #,##0.00');

        //FORMULE
        $sheet->setCellValue('C4', '=SUM(C2:C3)');
        //Ottenere il risultato di una formula
        //$sheet->getCell("C4")->getCalculatedValue();
        //Grassetto
        $sheet->getStyle('C4')->getFont()->setBold(true);

        $sheet->setCellValueByColumnAndRow(0, 10, 'CELLE UNITE');

        //Celle unite
        $sheet->mergeCells('A'.'10'.':C'.'10');

        //Wrap text
        $sheet->setCellValueByColumnAndRow(0, 11, 'TESTO A CAPO');
        $sheet->getStyle('A11:B11')->getAlignment()->setWrapText(true);

        $sheet2 = $objPHPExcel->createSheet();
        $sheet2->setTitle('SecondoFoglio');

        //Scrittura su file
        //Si crea un oggetto
        $objWriter = new \PHPExcel_Writer_Excel5($objPHPExcel);
        $todaydate = date('d-m-y');
        //$todaydate = $todaydate . '-' . date("H-i-s");
        $filename = 'nomefile';
        $filename = $filename.'-'.$todaydate;
        $filename = $filename.'.xls';
        $filename = sys_get_temp_dir().DIRECTORY_SEPARATOR.$filename;

        if (file_exists($filename)) {
            unlink($filename);
        }

        $objWriter->save($filename);

        $response = new Response();

        $response->headers->set('Content-Type', 'text/csv');
        $response->headers->set('Content-Disposition', 'attachment;filename="'.basename($filename).'"');

        $response->setContent(file_get_contents($filename));

        //Per avere disponibile al download il file excel scommentare
        //return $response; e commentare return $this->render('DemoBundle:Demo:output.html.twig');
        //return $response;
        return $this->render('DemoBundle:Demo:output.html.twig');
    }

    public function excelqueryToExcelAction()
    {
        set_time_limit(960);
        ini_set('memory_limit', '2048M');

        $cacheMethod = \PHPExcel_CachedObjectStorageFactory::cache_to_phpTemp;
        $cacheSettings = array('memoryCacheSize' => '8MB');
        \PHPExcel_Settings::setCacheStorageMethod($cacheMethod, $cacheSettings);

        //Creare un nuovo file
        $objPHPExcel = new \PHPExcel();

        $objPHPExcel->setActiveSheetIndex(0);

        // Set properties
        $objPHPExcel->getProperties()->setCreator('Comune di Firenze');
        $objPHPExcel->getProperties()->setLastModifiedBy('Comune di Firenze');
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->setTitle('Export');

        $queryObj = $this->get('oracle_manager');
        $sql = "SELECT * FROM ALL_TABLES WHERE OWNER = 'P00' AND ROWNUM < 30"; // AND ROWNUM < 30
        //$sql = "SELECT OWNER,TABLE_NAME,TABLESPACE_NAME,CLUSTER_NAME,IOT_NAME,STATUS,PCT_FREE,PCT_USED,INI_TRANS,MAX_TRANS,INITIAL_EXTENT
        //FROM ALL_TABLES
        //WHERE OWNER = 'P00' AND ROWNUM < 30";

        $queryObj->executeSelectQuery($sql, false);
        $resultset = $queryObj->getResultset();
        $numcol = count($resultset);
        // Si imposta il font
        //Times new romans
        $sheet->getDefaultStyle()->getFont()->setName('Verdana');
        $sheet->getDefaultStyle()->getFont()->setSize(12);

        //Si imposta il colore dello sfondo delle celle
        //Colore header
        $style_header = array(
            'fill' => array(
                'type' => \PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => 'CC99FF'),
            ),
            'font' => array(
                'bold' => false,
                'color' => array('rgb' => '000000'),
            ),
        );
        $sheet->getStyle('A1:'.\PHPExcel_Cell::stringFromColumnIndex($numcol - 1).'1')->applyFromArray($style_header);

        $col = 0;
        foreach ($resultset as $key => $rows) {
            $sheet->setCellValueByColumnAndRow($col, 1, $key);
            $row = 1;
            foreach ($rows as $value) {
                $row = $row + 1;
                $sheet->setCellValueByColumnAndRow($col, $row, $value);
            }
            $col = $col + 1;
        }

        //Scrittura su file
        //Si crea un oggetto
        $objWriter = new \PHPExcel_Writer_Excel5($objPHPExcel);
        $todaydate = date('d-m-y');
        //$todaydate = $todaydate . '-' . date("H-i-s");
        $filename = 'export';
        $filename = $filename.'-'.$todaydate;
        $filename = $filename.'.xls';
        $filename = sys_get_temp_dir().DIRECTORY_SEPARATOR.$filename;

        if (file_exists($filename)) {
            unlink($filename);
        }

        $objWriter->save($filename);

        $response = new Response();

        $response->headers->set('Content-Type', 'text/csv');
        $response->headers->set('Content-Disposition', 'attachment;filename="'.basename($filename).'"');

        $response->setContent(file_get_contents($filename));
        //Per avere disponibile al download il file excel scommentare
        //return $response; e commentare return $this->render('DemoBundle:Demo:output.html.twig');
        //return $response;
        return $this->render('DemoBundle:Demo:output.html.twig');
    }
}
