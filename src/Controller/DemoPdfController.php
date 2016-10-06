<?php

namespace Fi\DemoBundle\Controller;

use Symfony\Bundle\FrameworkBundle\Controller\Controller;
use Symfony\Component\HttpFoundation\Response;
use Symfony\Component\HttpFoundation\Request;
use PhpOffice\PhpWord\PhpWord;
use Fi\ImapBundle\DependencyInjection\ImapMailbox;
use Symfony\Component\HttpKernel\Kernel;

class DemoPdfController extends Controller {

    public function docx2PdfAction(Request $request) {
        //Aprire un file esistente
        $templatefilepath = $this->get('kernel')->getRootDir() . "/tmp/doc/attestato.docx";

        $doc = new PHPWord();
        $document = $doc->loadTemplate($templatefilepath);
        $document->setValue('MATRICOLA', 59495);
        $document->setValue('NOMINATIVO', "Andrea Manzi");
        $document->setValue('TITOLO', "Corso PHP Base");
        $document->setValue('DAL', "01/02/2015");
        $document->setValue('AL', "01/03/2015");
        $document->setValue('DURATA', 60);
        $document->setValue('ORE', 58);
        $document->setValue('AGENZIA', "IDI");
        $document->setValue('FORMATORE', "Francesco Leoncino");
        $document->setValue('DATA', date('d/m/Y'));

        $filename = "attestatogenerato";
        $fileattestato = $this->get('kernel')->getRootDir() . "/tmp/doc/" . $filename . ".docx";
        $fs = new \Symfony\Component\Filesystem\Filesystem();
        if ($fs->exists($fileattestato)) {
            $fs->remove($fileattestato);
        }

        $document->saveAs($fileattestato);

        if (!($fs->exists($fileattestato))) {
            return new Response("Impossibile creare il file " . $fileattestato);
        }

        $outdir = $this->get('kernel')->getRootDir() . "/tmp/pdf/";
        /* @var $fs \Symfony\Component\Filesystem\Filesystem */
        $fs->mkdir($outdir, 0777);

        if (!($fs->exists($outdir))) {
            return new Response("Impossibile creare la cartella " . $outdir);
        }

        $libreofficePath = "/usr/bin/libreoffice";

        $convertcmd = $libreofficePath . " --headless --convert-to pdf " . $fileattestato . " --outdir " . $outdir;
        /* @var $process \Symfony\Component\Process\Process */
        $process = new \Symfony\Component\Process\Process($convertcmd);

        //Per poter generare il file tramite libreoffice è necessario impostare la variabile env HOME per apache
        $process->setEnv(array("HOME" => "/tmp"));

        $process->run();

        //Si presume esista libreoffice, quindi controllare che sia installato 
        //perchè non potendo chiedere la file_exists per problemi di privilegi sul server
        if (!$process->isSuccessful()) {
            return new Response($process->getErrorOutput());
        } else {
            //echo $process->getOutput();exit;
            $pdf = $outdir . $filename . ".pdf";
            if ($fs->exists($pdf)) {
                $response = new Response();

                $response->headers->set('Content-Type', 'application/pdf');
                $response->headers->set('Content-Disposition', 'attachment;filename="' . basename($pdf) . '"');
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
                return new Response("Il server non e' stato in grado di generare il file pdf");
            }
        }
    }

}
