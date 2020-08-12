<?php


namespace App\Isotope\Document;

use Contao\CoreBundle\Exception\ResponseException;
use Contao\PageModel;
use Contao\System;
use DateTime;
use Isotope\Frontend;
use Isotope\Interfaces\IsotopeDocument;
use Isotope\Interfaces\IsotopeProductCollection;
use Isotope\Isotope;
use Isotope\Model\Document;
use PhpOffice\PhpWord\Settings;
use PhpOffice\PhpWord\TemplateProcessor;
use Symfony\Component\HttpFoundation\BinaryFileResponse;
use Symfony\Component\HttpFoundation\ResponseHeaderBag;

class InvoiceDoxc extends Document implements IsotopeDocument
{

    /**
     * {@inheritdoc}
     */
    public function outputToBrowser(IsotopeProductCollection $objCollection)
    {
        $this->prepareEnvironment($objCollection);

        $tokens   = $this->prepareCollectionTokens($objCollection);
        $document = $this->generateDocument($objCollection, $tokens);
        $tmpPath  = sprintf('%s/system/tmp/isotope/%s.docx', TL_ROOT, uniqid('', true));
        $fileName = $this->prepareFileName($this->fileTitle, $tokens) . '.docx';

        $document->saveAs($tmpPath);

        $response = new BinaryFileResponse($tmpPath);
        $response->setContentDisposition(ResponseHeaderBag::DISPOSITION_ATTACHMENT, $fileName);

        throw new ResponseException($response);
    }

    /**
     * {@inheritdoc}
     */
    public function outputToFile(IsotopeProductCollection $objCollection, $strDirectoryPath)
    {
        $this->prepareEnvironment($objCollection);

        $tokens   = $this->prepareCollectionTokens($objCollection);
        $document = $this->generateDocument($objCollection, $tokens);
        $path     = $this->prepareFileName($this->fileTitle, $tokens, $strDirectoryPath) . '.docx';

        $document->saveAs($path);

        return $path;
    }

    protected function generateDocument(IsotopeProductCollection $order, array $tokens): TemplateProcessor
    {
        Settings::setOutputEscapingEnabled(true);

        $templatePath = System::getContainer()->getParameter('app.isotope_invoice_doxc');

        $dateFormat = $GLOBALS['TL_CONFIG']['dateFormat'];
        $orderDate  = $order->getLockTime();
        $orderTotal = $order->getTotal();
        $dueDate    = (new DateTime('@' . $orderDate))->modify('+28 days');

        $templateProcessor = new TemplateProcessor($templatePath);

        // Variables on different parts of document
        $templateProcessor->setValue('orderTotal', $this->formatPrice($orderTotal));
        $templateProcessor->setValue('orderDate', date($dateFormat, $orderDate));
        $templateProcessor->setValue('dueDate', $dueDate->format($dateFormat));

        // Billing address
        if (null !== $billingAddress = $order->getBillingAddress()) {
            $templateProcessor->setValue('billingCompany', $billingAddress->company);
            $templateProcessor->setValue('billingSalutation', $billingAddress->salutation);
            $templateProcessor->setValue('billingFirstname', $billingAddress->firstname);
            $templateProcessor->setValue('billingLastname', $billingAddress->lastname);
            $templateProcessor->setValue('billingStreet', $billingAddress->street_1);
            $templateProcessor->setValue('billingPostal', $billingAddress->postal);
            $templateProcessor->setValue('billingCity', $billingAddress->city);
        }

        // Order items
        $templateProcessor->cloneRow('itemRowName', $order->countItems());

        $i = 0;
        foreach ($order->getItems() as $item) {
            ++$i;
            $templateProcessor->setValue('itemRowName#' . $i, $item->getName());
            $templateProcessor->setValue('itemRowQuantity#' . $i, $item->quantity);
            $templateProcessor->setValue('itemRowPrice#' . $i, $this->formatPrice($item->getPrice()));
            $templateProcessor->setValue('itemRowPriceTotal#' . $i, $this->formatPrice($item->getTotalPrice()));
        }

        return $templateProcessor;
    }

    protected function prepareEnvironment(IsotopeProductCollection $objCollection): void
    {
        /** @var PageModel&\PageModel $objPage */
        global $objPage;

        if (!\is_object($objPage) && $objCollection->pageId > 0) {
            $objPage = PageModel::findWithDetails($objCollection->pageId);
            $objPage = Frontend::loadPageConfig($objPage);

            System::loadLanguageFile('default', $GLOBALS['TL_LANGUAGE'], true);
        }
    }

    private function formatPrice(float $price)
    {
        return Isotope::formatPriceWithCurrency($price, false);
    }
}
