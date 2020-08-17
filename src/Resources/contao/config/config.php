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

use Isotope\Model\Document;
use Terminal42\IsotopeOoxmlDocument\Isotope\Document\WordTemplate;

Document::registerModelType('word_template', WordTemplate::class);
