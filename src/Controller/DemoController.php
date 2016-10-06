<?php

namespace Fi\DemoBundle\Controller;

use Symfony\Bundle\FrameworkBundle\Controller\Controller;
use Symfony\Component\HttpFoundation\Response;
use Symfony\Component\HttpFoundation\Request;
use Symfony\Component\HttpKernel\Kernel;

class DemoController extends Controller
{
    public function indexAction()
    {
        //Prova
        return $this->render('DemoBundle:Demo:index.html.twig');
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
}
