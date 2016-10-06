<?php

namespace Fi\DemoBundle\Controller;

use Symfony\Bundle\FrameworkBundle\Controller\Controller;
use Fi\ImapBundle\DependencyInjection\ImapMailbox;
use Symfony\Component\HttpKernel\Kernel;

class DemoMailController extends Controller
{
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

        $mailbox = new ImapMailbox($indirizzomail, $utentemail, $passwordmail, 'UTF-8');

        $arraymessaggi = array();
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
}
