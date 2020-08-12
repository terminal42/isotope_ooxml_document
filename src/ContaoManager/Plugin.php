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

namespace Terminal42\IsotopeOoxmlDocument\ContaoManager;

use Contao\CoreBundle\ContaoCoreBundle;
use Contao\ManagerPlugin\Bundle\BundlePluginInterface;
use Contao\ManagerPlugin\Bundle\Config\BundleConfig;
use Contao\ManagerPlugin\Bundle\Parser\ParserInterface;
use Terminal42\IsotopeOoxmlDocument\Terminal42IsotopeOoxmlDocumentBundle;

final class Plugin implements BundlePluginInterface
{
    public function getBundles(ParserInterface $parser): array
    {
        return [
            BundleConfig::create(Terminal42IsotopeOoxmlDocumentBundle::class)
                ->setLoadAfter([ContaoCoreBundle::class, 'isotope']),
        ];
    }
}
