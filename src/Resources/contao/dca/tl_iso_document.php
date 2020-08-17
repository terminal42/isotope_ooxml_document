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

$GLOBALS['TL_DCA']['tl_iso_document']['palettes']['word_template'] =
    '{type_legend},name,type;{config_legend},documentTitle,fileTitle;{template_legend},wordDocumentTpl,orderCollectionBy';

$GLOBALS['TL_DCA']['tl_iso_document']['fields']['wordDocumentTpl'] = [
    'inputType' => 'fileTree',
    'exclude'   => true,
    'eval'      => [
        'mandatory'  => true,
        'fieldType'  => 'radio',
        'multiple'   => false,
        'files'      => true,
        'filesOnly'  => true,
        'extensions' => 'doc,docx',
        'tl_class'   => 'clr',
    ],
    'sql'       => 'binary(16) NULL',
];
