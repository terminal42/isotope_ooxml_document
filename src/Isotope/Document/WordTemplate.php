<?php

declare(strict_types=1);

/*
 * OOXML Office Documents bundle for Isotope eCommerce.
 *
 * @copyright  Copyright (c) 2020, terminal42 gmbh
 * @author     terminal42 <https://terminal42.ch>
 * @license    MIT
 * @link       http://github.com/terminal42/isotope_ooxml_document
 */

namespace Terminal42\IsotopeOoxmlDocument\Isotope\Document;

use Contao\CoreBundle\Exception\ResponseException;
use Contao\FilesModel;
use Contao\PageModel;
use Contao\System;
use Isotope\Frontend;
use Isotope\Interfaces\IsotopeDocument;
use Isotope\Interfaces\IsotopeProductCollection;
use Isotope\Interfaces\IsotopePurchasableCollection;
use Isotope\Isotope;
use Isotope\Model\Document;
use Isotope\Model\ProductCollection;
use PhpOffice\PhpWord\Exception\Exception as WordException;
use PhpOffice\PhpWord\Settings;
use PhpOffice\PhpWord\TemplateProcessor;
use Symfony\Component\HttpFoundation\BinaryFileResponse;
use Symfony\Component\HttpFoundation\ResponseHeaderBag;

/**
 * This document model takes a docx-file as input,
 * uses the template processor to replace variables in the template,
 * and outputs the word document (untouched file unless variables processing).
 */
class WordTemplate extends Document implements IsotopeDocument
{
    /**
     * {@inheritdoc}
     */
    public function outputToBrowser(IsotopeProductCollection $objCollection)
    {
        $this->prepareEnvironment($objCollection);

        $tokens    = $this->prepareCollectionTokens($objCollection);
        $variables = $this->generateDocumentVariables($objCollection);
        $document  = $this->generateDocument($objCollection, $variables);
        $fileName  = $this->prepareFileName($this->fileTitle, $tokens).'.docx';
        $tmpPath   = $document->save();

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

        $tokens    = $this->prepareCollectionTokens($objCollection);
        $variables = $this->generateDocumentVariables($objCollection);
        $document  = $this->generateDocument($objCollection, $variables);
        $path      = $this->prepareFileName($this->fileTitle, $tokens, $strDirectoryPath).'.docx';

        $document->saveAs($path);

        return $path;
    }

    /**
     * Use the template processor to replace variables in the word document.
     *
     * Variables are embedded like `${variable_name}` in the word document. Variables MUST be single line,
     * multiline content cannot be processed.
     *
     * For the collection items, we use the `cloneRow("item_name", n)` method. The word template should
     * provide the following table:  | Quantity         | Product      | Price         | Total Price         |
     *                               | ${item_quantity} | ${item_name} | {$item_price} | ${item_total_price} |
     *                               |                  |              | Total         | ${order_total}      |
     *
     * As you can see, the second row will be duplicated because it holds the "item_name" variable.
     * It is required to have the "item_name" column in the table though.
     * The table can be extended with a placeholder row for the surcharges (row must contain `${surcharge_label}`).
     * You may then want to place a row with `${order_subtotal}` in between.
     */
    protected function generateDocument(IsotopeProductCollection $order, array $variables): TemplateProcessor
    {
        $pageModel  = PageModel::findWithDetails($order->page_id);
        $dateFormat = $pageModel->dateFormat ?: $GLOBALS['TL_CONFIG']['dateFormat'];

        Settings::setOutputEscapingEnabled(true);

        $templateFile = FilesModel::findByPk($this->wordDocumentTpl);
        if (null === $templateFile) {
            throw new \LogicException('Could not find word document template. Make sure to have a word document assigned in the template configuration.');
        }

        $templateProcessor = new TemplateProcessor(TL_ROOT.'/'.$templateFile->path);

        // Set the variables to replace in the template.
        foreach ($variables as $search => $replace) {
            $templateProcessor->setValue($search, $replace);
        }

        $templateProcessor->setValue('order_date', date($dateFormat, (int) $order->getLockTime()));

        // Now: process the collection items.
        $this->processItems($templateProcessor, $order);

        // Same procedure for the surcharges (shipping fee, taxes etc.)
        $this->processSurcharges($order, $templateProcessor);

        // If you want to add your own variables to processed word template, use the getNotificationTokens hook.
        // Might add an event here later, to allow further modification of the template processing.

        return $templateProcessor;
    }

    private function generateDocumentVariables(IsotopeProductCollection $collection): array
    {
        if (!$collection instanceof IsotopePurchasableCollection) {
            return [];
        }

        // Since we don't pass a notification id, the tokens "cart_text" and "cart_html"
        // will not be generated, which we cannot process anyway.
        return $collection->getNotificationTokens(0);
    }

    private function processItems(TemplateProcessor $templateProcessor, IsotopeProductCollection $order): void
    {
        $sortCallback = ProductCollection::getItemsSortingCallable($this->orderCollectionBy);
        try {
            $templateProcessor->cloneRow('item_name', $order->countItems());

            $i = 0;
            foreach ($order->getItems($sortCallback) as $item) {
                ++$i;

                $templateProcessor->setValue(sprintf('item_%s#%d', 'name', $i), html_entity_decode($item->getName()));
                $templateProcessor->setValue(sprintf('item_%s#%d', 'quantity', $i), $item->quantity);
                $templateProcessor->setValue(sprintf('item_%s#%d', 'price', $i), $this->formatPrice($item->getPrice()));
                $templateProcessor->setValue(sprintf('item_%s#%d', 'total_price', $i), $this->formatPrice($item->getTotalPrice()));
            }
        } catch (WordException $e) {
        }
    }

    private function processSurcharges(IsotopeProductCollection $order, TemplateProcessor $templateProcessor): void
    {
        $surcharges = $order->getSurcharges();
        try {
            $templateProcessor->cloneRow('surcharge_label', \count($surcharges));

            $i = 0;
            foreach ($surcharges as $surcharge) {
                ++$i;

                $templateProcessor->setValue(sprintf('surcharge_%s#%d', 'label', $i), $surcharge->label);
                $templateProcessor->setValue(sprintf('surcharge_%s#%d', 'price', $i), $this->formatPrice($surcharge->price));
                $templateProcessor->setValue(sprintf('surcharge_%s#%d', 'total_price', $i), $this->formatPrice($surcharge->total_price));
            }
        } catch (WordException $e) {
        }
    }

    private function prepareEnvironment(IsotopeProductCollection $collection): void
    {
        /* @var PageModel|\PageModel $objPage */
        global $objPage;

        if (!\is_object($objPage) && $collection->pageId > 0) {
            $objPage = PageModel::findWithDetails($collection->pageId);
            $objPage = Frontend::loadPageConfig($objPage);

            System::loadLanguageFile('default', $GLOBALS['TL_LANGUAGE'], true);
        }
    }

    private function formatPrice($price)
    {
        return Isotope::formatPriceWithCurrency($price, false);
    }
}
