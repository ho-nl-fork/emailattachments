<?php
/**
 * @author     Kristof Ringleff
 * @package    Fooman_EmailAttachments
 * @copyright  Copyright (c) 2009 Fooman Limited (http://www.fooman.co.nz)
 *
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */
class Fooman_EmailAttachments_Model_Observer
{
    const XML_PATH_ORDER_PACKINGSLIP_TEMPLATE   = 'sales_email/order/shipment_template';
    const XML_PATH_EMAIL_COPY_TO                = 'sales_email/order/shipment_to';
    const KEY_PACKING_SLIP_PROCESSED            = 'emailattachments-packingslip-processed';

    protected function _fixUnsavedComments($object)
    {
        if (Mage::app()->getRequest()->getPost('comment')) {
            $object->save();
        }
    }

    /**
     * observe core_block_abstract_prepare_layout_after to add a Print Orders
     * massaction to the actions dropdown menu
     *
     * @param   Varien_Event_Observer   $observer
     */
    public function addbutton(Varien_Event_Observer $observer)
    {
        $block = $observer->getEvent()->getBlock();
        //add button to dropdown
        if ($block instanceof Mage_Adminhtml_Block_Widget_Grid_Massaction
            || $block instanceof
            Enterprise_SalesArchive_Block_Adminhtml_Sales_Order_Grid_Massaction
        ) {
            if ($block->getRequest()->getControllerName() == 'sales_order'
                || $block->getRequest()->getControllerName() == 'adminhtml_sales_order'
                || $block->getRequest()->getControllerName() == 'sales_archive'
                || $block->getRequest()->getControllerName() == 'orderspro_order'
            ) {
                $block->addItem(
                    'fooman_pdforders_order', array(
                        'label' => Mage::helper('emailattachments')->__('Print Orders'),
                        'url'   => Mage::helper('adminhtml')->getUrl(
                            'adminhtml/EmailAttachments_order/pdforders',
                            Mage::app()->getStore()->isCurrentlySecure() ? array('_secure' => 1) : array()
                        ),
                    )
                );
            }
        }
        //add button to single order view
        if ($block instanceof Mage_Adminhtml_Block_Sales_Order_View) {
            Mage::helper('emailattachments')->addButton($block);
        }
    }

    /**
     * Listen to order email send event to attach pdf's and agreements.
     *
     * @event   fooman_emailattachments_before_send_order
     * @param   Varien_Event_Observer   $observer
     *
     * @return  Fooman_EmailAttachments_Model_Observer
     */
    public function beforeSendOrder(Varien_Event_Observer $observer)
    {
        $update         = $observer->getEvent()->getUpdate();
        $mailTemplate   = $observer->getEvent()->getTemplate();
        $order          = $observer->getEvent()->getObject();
        $storeId        = $order->getStoreId();
        $configPath     = $update ? 'order_comment' : 'order';

        /** @var Fooman_EmailAttachments_Helper_Data $helper */
        $helper = Mage::helper('emailattachments');

        if (Mage::getStoreConfig('sales_email/' . $configPath . '/attachpdf', $storeId)) {
            $pdf = $helper->getPdf($order, 'order');

            $helper->addAttachment(
                $pdf, $mailTemplate, $this->getOrderAttachmentName($order)
            );
        }

        if (Mage::getStoreConfig('sales_email/' . $configPath . '/attachagreement', $storeId)) {
            $helper->addAgreements($order->getStoreId(), $mailTemplate);
        }

        for ($i = 0; $i <= 5; $i++) {
            $fileAttachment = Mage::getStoreConfig('sales_email/'.$configPath.'/attachfile_'.$i, $storeId);
            if ($fileAttachment) {
                $helper->addFileAttachment($fileAttachment, $mailTemplate);
            }
        }

        return $this;
    }

    public function getOrderAttachmentName($order)
    {
        return Mage::helper('emailattachments')->getOrderAttachmentName($order);
    }

    /**
     * listen to order email send event to send packing slip
     *
     * @param $observer
     */
    public function sendPackingSlip($observer)
    {
        if (!Mage::registry(self::KEY_PACKING_SLIP_PROCESSED)) {
            Mage::register(self::KEY_PACKING_SLIP_PROCESSED, true);
        } else {
            //only process this once
            return;
        }
        $update = $observer->getEvent()->getUpdate();
        $mailTemplate = Mage::getModel('core/email_template');
        $order = $observer->getEvent()->getObject();
        $configPath = $update ? 'order_comment' : 'order';
        $storeId = $order->getStoreId();
        $emails = Mage::helper('emailattachments')->getEmails(self::XML_PATH_EMAIL_COPY_TO, $storeId);

        if ($emails && Mage::getStoreConfig('sales_email/' . $configPath . '/sendpackingslip', $storeId)) {
            $template = Mage::getStoreConfig(self::XML_PATH_ORDER_PACKINGSLIP_TEMPLATE, $storeId);
            $pdf = Mage::getModel('sales/order_pdf_shipment')->getPdf(array(), array($order->getId()));
            Mage::helper('emailattachments')->addAttachment(
                $pdf, $mailTemplate, Mage::helper('sales')->__('Shipment') . "_" . $order->getIncrementId()
            );
            foreach ($emails as $email) {
                $mailTemplate->setDesignConfig(array('area' => 'frontend', 'store' => $storeId))
                    ->sendTransactional(
                        $template,
                        Mage::getStoreConfig(
                            Mage_Sales_Model_Order_Shipment::XML_PATH_EMAIL_IDENTITY, $storeId
                        ),
                        $email,
                        '',
                        array('order' => $order)
                    );
            }
        }


    }

    /**
     * Listen to invoice email send event to attach pdf's and agreements.
     *
     * @event   fooman_emailattachments_before_send_invoice
     * @param   Varien_Event_Observer $observer
     *
     * @return  Fooman_EmailAttachments_Model_Observer
     */
    public function beforeSendInvoice(Varien_Event_Observer $observer)
    {
        $update         = $observer->getEvent()->getUpdate();
        $mailTemplate   = $observer->getEvent()->getTemplate();
        $invoice        = $observer->getEvent()->getObject();
        $storeId        = $invoice->getStoreId();
        $configPath     = $update ? 'invoice_comment' : 'invoice';

        /** @var Fooman_EmailAttachments_Helper_Data $helper */
        $helper = Mage::helper('emailattachments');

        if (Mage::getStoreConfig('sales_email/' . $configPath . '/attachpdf', $storeId)) {
            $this->_fixUnsavedComments($invoice);

            $pdf = $helper->getPdf($invoice, 'invoice');

            $helper->addAttachment(
                $pdf, $mailTemplate, $this->getInvoiceAttachmentName($invoice)
            );
        }

        if (Mage::getStoreConfig('sales_email/' . $configPath . '/attachagreement', $storeId)) {
            $helper->addAgreements($storeId, $mailTemplate);
        }

        for ($i = 0; $i <= 5; $i++) {
            $fileAttachment = Mage::getStoreConfig('sales_email/'.$configPath.'/attachfile_'.$i, $storeId);
            if ($fileAttachment) {
                $helper->addFileAttachment($fileAttachment, $mailTemplate);
            }
        }
    }

    public function getInvoiceAttachmentName($invoice)
    {
        return Mage::helper('emailattachments')->getInvoiceAttachmentName($invoice);
    }

    /**
     * Listen to shipment email send event to attach pdf's and agreements.
     *
     * @event   fooman_emailattachments_before_send_shipment
     * @param   Varien_Event_Observer   $observer
     *
     * @return  Fooman_EmailAttachments_Model_Observer
     */
    public function beforeSendShipment(Varien_Event_Observer $observer)
    {
        $update         = $observer->getEvent()->getUpdate();
        $mailTemplate   = $observer->getEvent()->getTemplate();
        $shipment       = $observer->getEvent()->getObject();
        $storeId        = $shipment->getStoreId();
        $configPath     = $update ? 'shipment_comment' : 'shipment';

        /** @var Fooman_EmailAttachments_Helper_Data $helper */
        $helper = Mage::helper('emailattachments');

        if (Mage::getStoreConfig('sales_email/' . $configPath . '/attachpdf', $storeId)) {
            $this->_fixUnsavedComments($shipment);

            $pdf = $helper->getPdf($shipment, 'shipment');

            $helper->addAttachment(
                $pdf, $mailTemplate, $this->getShipmentAttachmentName($shipment)
            );
        }

        if (Mage::getStoreConfig('sales_email/' . $configPath . '/attachagreement', $storeId)) {
            $helper->addAgreements($storeId, $mailTemplate);
        }

        for ($i = 0; $i <= 5; $i++) {
            $fileAttachment = Mage::getStoreConfig('sales_email/'.$configPath.'/attachfile_'.$i, $storeId);
            if ($fileAttachment) {
                $helper->addFileAttachment($fileAttachment, $mailTemplate);
            }
        }

        return $this;
    }

    public function getShipmentAttachmentName($shipment)
    {
        return Mage::helper('emailattachments')->getShipmentAttachmentName($shipment);
    }

    /**
     * Listen to creditmemo email send event to attach pdf's and agreements.
     *
     * @event   fooman_emailattachments_before_send_creditmemo
     * @param   Varien_Event_Observer   $observer
     *
     * @return  Fooman_EmailAttachments_Model_Observer
     */
    public function beforeSendCreditmemo(Varien_Event_Observer $observer)
    {
        $update         = $observer->getEvent()->getUpdate();
        $mailTemplate   = $observer->getEvent()->getTemplate();
        $creditmemo     = $observer->getEvent()->getObject();
        $storeId        = $creditmemo->getStoreId();
        $configPath     = $update ? 'creditmemo_comment' : 'creditmemo';

        /** @var Fooman_EmailAttachments_Helper_Data $helper */
        $helper = Mage::helper('emailattachments');

        if (Mage::getStoreConfig('sales_email/' . $configPath . '/attachpdf', $storeId)) {
            $this->_fixUnsavedComments($creditmemo);

            $pdf = $helper->getPdf($creditmemo, 'creditmemo');

            $helper->addAttachment(
                $pdf, $mailTemplate, $this->getCreditmemoAttachmentName($creditmemo)
            );
        }

        if (Mage::getStoreConfig('sales_email/' . $configPath . '/attachagreement', $storeId)) {
            $helper->addAgreements($storeId, $mailTemplate);
        }

        for ($i = 0; $i <= 5; $i++) {
            $fileAttachment = Mage::getStoreConfig('sales_email/'.$configPath.'/attachfile_'.$i, $storeId);
            if ($fileAttachment) {
                $helper->addFileAttachment($fileAttachment, $mailTemplate);
            }
        }

        return $this;
    }

    public function getCreditmemoAttachmentName($creditmemo)
    {
        return Mage::helper('emailattachments')->getCreditmemoAttachmentName($creditmemo);
    }

    protected function _processSendQueueEvent($mailer, $message)
    {
        if ($message->getEntityType() == 'order' && !$message->getProcessedAt()) {
            $order      = Mage::getModel('sales/order')->load($message->getEntityId());
            $storeId    = $order->getStoreId();
            $update     = $message->getEventType() == 'update_order';
            $configPath = $update ? 'order_comment' : 'order';

            /** @var Fooman_EmailAttachments_Helper_Data $helper */
            $helper = Mage::helper('emailattachments');

            if (Mage::getStoreConfig('sales_email/' . $configPath . '/attachpdf', $storeId)) {
                $pdf = $helper->getPdf($order, 'order');

                $helper->addAttachment(
                    $pdf, $mailer, $this->getOrderAttachmentName($order)
                );
            }

            if (Mage::getStoreConfig('sales_email/' . $configPath . '/attachagreement', $storeId)) {
                $helper->addAgreements($order->getStoreId(), $mailer);
            }

            for ($i = 0; $i <= 5; $i++) {
                $fileAttachment = Mage::getStoreConfig('sales_email/'.$configPath.'/attachfile_'.$i, $storeId);
                if ($fileAttachment) {
                   $helper->addFileAttachment($fileAttachment, $mailer);
                }
            }
        }
    }

    /**
     * listen to order email send event from queue to attach pdfs and agreements
     *
     * @param   $observer
     */
    public function beforeSendQueuedOrder($observer)
    {
        $mailer = $observer->getEvent()->getMailer()
            ? $observer->getEvent()->getMailer()
            : $observer->getEvent()->getMail();
        $message = $observer->getEvent()->getMessage();

        $this->_processSendQueueEvent($mailer, $message);
    }
}
