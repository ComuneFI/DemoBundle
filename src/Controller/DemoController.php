<?php

namespace Fi\DemoBundle\Controller;

use Symfony\Bundle\FrameworkBundle\Controller\Controller;
use Symfony\Component\HttpFoundation\Response;
use Symfony\Component\HttpFoundation\Request;
use PhpOffice\PhpWord\PhpWord;
use Fi\ImapBundle\DependencyInjection\ImapMailbox;
use Symfony\Component\HttpKernel\Kernel;

class DemoController extends Controller
{
    public function indexAction()
    {
        //Prova
        return $this->render('DemoBundle:Demo:index.html.twig');
    }

    public function doctrineInsertAction()
    {
        /* @var $em \Doctrine\ORM\EntityManager */
        $em = $this->get('doctrine')->getManager();

        $nuovoOperatore = new \Fi\CoreBundle\Entity\operatori();
        $nuovoOperatore->setOperatore('CognomeNome');
        $nuovoOperatore->setUsername('Dxxxxx');
        $ruolo = $em->getRepository('FiCoreBundle:ruoli')->find(1);
        $nuovoOperatore->setRuoli($ruolo);
        //Togliere il commento alla riga successiva per rendere definitiva la modifica sul database
        //$em->persist($nuovoOperatore);
        $em->flush();

        return $this->render('DemoBundle:Demo:output.html.twig');
    }

    public function doctrineDeleteAction()
    {
        /* @var $em \Doctrine\ORM\EntityManager */
        $em = $this->get('doctrine')->getManager();

        $qb = $em->createQueryBuilder();
        $qb->delete('FiCoreBundle:operatori', 's');
        $qb->andWhere($qb->expr()->eq('s.id', ':idOperatore'));
        $qb->setParameter(':idOperatore', 12);
        $qb->getQuery()->execute();

        return $this->render('DemoBundle:Demo:output.html.twig');
    }

    public function doctrineUpdateAction()
    {
        /* @var $em \Doctrine\ORM\EntityManager */
        $em = $this->get('doctrine')->getManager();

        $qb = $em->createQueryBuilder();
        $qb->update('FiCoreBundle:operatori', 's');
        $qb->set('s.username', ':newValue');
        $qb->andWhere($qb->expr()->eq('s.id', ':idOperatore'));
        $qb->setParameter(':idOperatore', 3);
        $qb->setParameter(':newValue', 'DXxXxXx');
        $qb->getQuery()->execute();

        return $this->render('DemoBundle:Demo:output.html.twig');
    }

    public function doctrineSelectAction()
    {
        /* @var $em \Doctrine\ORM\EntityManager */
        $em = $this->get('doctrine')->getManager();

        /* @var $qb \Doctrine\ORM\QueryBuilder */
        $qb = $em->createQueryBuilder();
        $qb->select(array('a'));
        $qb->from('FiCoreBundle:operatori', 'a');
        $qb->where('a.id = :idOperatore');
        $qb->setParameter('idOperatore', 1);
        //$qb->setFirstResult( $offset )
        //$qb->setMaxResults( $limit );
        $resultset = $qb->getQuery()->getResult();

        foreach ($resultset as $row) {
            /* ... per ogni elemento ... */
        }

        return $this->render('DemoBundle:Demo:output.html.twig');
    }

    public function uploadIndexAction()
    {
        return $this->render('DemoBundle:Demo:uploadfile.html.twig');
    }

    public function uploadSaveAction(Request $request)
    {
        //Cambiare la cartella $destinationFolder con quella in cui si desidera salvare il file uplodato
        $destinationFolder = $this->get('kernel')->getRootDir().'/tmp/';
        //Se esiste il file lo sovrascrive
        foreach ($request->files as $file) {
            $file->move($destinationFolder, $file->getClientOriginalName());
        }

        return new Response('OK');
    }

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

    public function inviomailAction()
    {
        $mittente = array('fakemittentemail@fakemail.com' => 'Fake Mittente');
        $destinatari = array();
        $cc = array();
        $bcc = array();
        $destinatari[] = 'fakemittentemail@fakemail.com';
        $cc[] = 'fakeccmail@fakemail.com';
        $bcc[] = 'fakebccmail@fakemail.com';

        $messaggio = \Swift_Message::newInstance()
                ->setSubject('Oggetto della mail')
                ->setFrom($mittente);
        $messaggio->setTo($destinatari);
        $messaggio->setCc($cc);
        $messaggio->setBcc($bcc);
        $messaggio->setBody("Corpo della mail\n");

        //Allegato
        $pathallegato = $this->get('kernel')->getRootDir().'/tmp/rec.xls';
        $messaggio->attach(\Swift_Attachment::fromPath($pathallegato));

        //ATTENZIONE! Questo comando invia la mail
        $this->get('mailer')->send($messaggio);

        return $this->render('DemoBundle:Demo:output.html.twig');
    }

    public function letturamailboxAction()
    {
        $indirizzomail = '{imap.comune.intranet:143/novalidate-cert}INBOX';
        $utentemail = 'imaptestlocale';
        $passwordmail = 'firenze1';

        $nummessaggiConsegna = 0;
        $mailbox = new ImapMailbox($indirizzomail, $utentemail, $passwordmail, 'UTF-8');

        $arraymessaggi = array();
        $mess = 0;
        $mailsIds = $mailbox->searchMailBox('ALL');
        if (!$mailsIds) {
            //Gestire come si vuole il fatto che non ci sono messaggi nella casella di posta
            throw new ImapMailboxException('Nessun messaggio trovato nella casella');
        } else {
            foreach ($mailsIds as $mailId) {
                $ok = true;
                try {
                    /* @var $mail \Fi\ImapBundle\DependencyInjection\IncomingMail */
                    $mail = $mailbox->getMail($mailId);

                    if (!$mail) {
                        $arraymessaggi[$mailId] = "** Errore parse headers del messaggio con ID $mailId";
                        $ok = false;
                    }
                } catch (Exception $ex) {
                    $arraymessaggi[$mailId] = "** Messaggio con caratteri errati - MailId $mailId ** Eccezione ".$ex->getTraceAsString();
                    $ok = false;
                }
                if ($ok === true) {
                    $arraymessaggi[$mailId]['id'] = $mail->id;
                    $arraymessaggi[$mailId]['subject'] = $mail->subject;
                    $arraymessaggi[$mailId]['bodytext'] = trim($mail->textPlain);
                    //$arraymessaggi[$mailId]["bodyhtml"] = trim($mail->textHtml);
                    $arraymessaggi[$mailId]['fromname'] = $mail->fromName;
                    $arraymessaggi[$mailId]['fromaddress'] = $mail->fromAddress;
                    $arraymessaggi[$mailId]['date'] = \DateTime::createFromFormat('Y-m-d H:i:s', $mail->date);
                    $arraymessaggi[$mailId]['replyto'] = $mail->replyTo;
                    $arraymessaggi[$mailId]['cc'] = $mail->cc;
                    $arraymessaggi[$mailId]['to'] = $mail->to;
                }
            }
        }

        return $this->render('DemoBundle:Demo:output.html.twig', array('extrainfo' => 'messaggi trovati:'.count($arraymessaggi)));
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
            $datanascita = \PHPExcel_Style_NumberFormat::toFormattedString($sheet->getCellByColumnAndRow($datanascitacol, $row)->getValue(), 'DD/MM/YYYY');

            //Leggere il valore del risultato di una formula
            $formula = $sheet->getCellByColumnAndRow($matricolacol, $row)->getCalculatedValue();

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

        //Per avere disponibile al download il file excel scommentare //return $response; e commentare return $this->render('DemoBundle:Demo:output.html.twig');
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
        //$sql = "SELECT OWNER,TABLE_NAME,TABLESPACE_NAME,CLUSTER_NAME,IOT_NAME,STATUS,PCT_FREE,PCT_USED,INI_TRANS,MAX_TRANS,INITIAL_EXTENT FROM ALL_TABLES WHERE OWNER = 'P00' AND ROWNUM < 30";

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
        //Per avere disponibile al download il file excel scommentare //return $response; e commentare return $this->render('DemoBundle:Demo:output.html.twig');
        //return $response;
        return $this->render('DemoBundle:Demo:output.html.twig');
    }

    public function docx2PdfAction(Request $request)
    {
        //Aprire un file esistente
        $templatefilepath = $this->get('kernel')->getRootDir().'/tmp/doc/attestato.docx';

        $doc = new PHPWord();
        $document = $doc->loadTemplate($templatefilepath);
        $document->setValue('MATRICOLA', 59495);
        $document->setValue('NOMINATIVO', 'Andrea Manzi');
        $document->setValue('TITOLO', 'Corso PHP Base');
        $document->setValue('DAL', '01/02/2015');
        $document->setValue('AL', '01/03/2015');
        $document->setValue('DURATA', 60);
        $document->setValue('ORE', 58);
        $document->setValue('AGENZIA', 'IDI');
        $document->setValue('FORMATORE', 'Francesco Leoncino');
        $document->setValue('DATA', date('d/m/Y'));

        $filename = 'attestatogenerato';
        $fileattestato = $this->get('kernel')->getRootDir().'/tmp/doc/'.$filename.'.docx';
        $fs = new \Symfony\Component\Filesystem\Filesystem();
        if ($fs->exists($fileattestato)) {
            $fs->remove($fileattestato);
        }

        $document->saveAs($fileattestato);

        if (!($fs->exists($fileattestato))) {
            echo 'Impossibile creare il file '.$fileattestato;
            exit;
        }

        $outdir = $this->get('kernel')->getRootDir().'/tmp/pdf/';
        /* @var $fs \Symfony\Component\Filesystem\Filesystem */
        $fs->mkdir($outdir, 0777);

        if (!($fs->exists($outdir))) {
            echo 'Impossibile creare la cartella '.$outdir;
            exit;
        }

        $libreofficePath = '/usr/bin/libreoffice';

        $convertcmd = $libreofficePath.' --headless --convert-to pdf '.$fileattestato.' --outdir '.$outdir;
        /* @var $process \Symfony\Component\Process\Process */
        $process = new \Symfony\Component\Process\Process($convertcmd);

        //Per poter generare il file tramite libreoffice è necessario impostare la variabile env HOME per apache
        $process->setEnv(array('HOME' => '/tmp'));

        $process->run();

        //Si presume esista libreoffice, quindi controllare che sia installato
        //perchè non potendo chiedere la file_exists per problemi di privilegi sul server
        if (!$process->isSuccessful()) {
            echo $process->getErrorOutput();
            exit;
        } else {
            //echo $process->getOutput();exit;
            $pdf = $outdir.$filename.'.pdf';
            if ($fs->exists($pdf)) {
                $response = new Response();

                $response->headers->set('Content-Type', 'application/pdf');
                $response->headers->set('Content-Disposition', 'attachment;filename="'.basename($pdf).'"');
                $response->setContent(file_get_contents($pdf));
                //Si cancellano i file docx e pdf tanto ormai è nell'header pronto per essere scaricato
                if ($fs->exists($fileattestato)) {
                    $fs->remove($fileattestato);
                }
                if ($fs->exists($pdf)) {
                    $fs->remove($pdf);
                }
                //Togliere commento a return $response; per scaricare il file
                //return $response;
                //Questo render solo per vedere il codice che sta dietro a questo controller
                return $this->render('DemoBundle:Demo:output.html.twig');
            } else {
                echo "Il server non e' stato in grado di generare il file pdf";
                exit;
            }
        }
    }
}
