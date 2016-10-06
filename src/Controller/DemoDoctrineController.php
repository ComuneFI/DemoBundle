<?php

namespace Fi\DemoBundle\Controller;

use Symfony\Bundle\FrameworkBundle\Controller\Controller;

class DemoDoctrineController extends Controller
{
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
            $operatoreid = $row->getId();
            echo $operatoreid."\n";
        }

        return $this->render('DemoBundle:Demo:output.html.twig');
    }
}
