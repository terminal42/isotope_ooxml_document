terminal42/isotope_ooxml_document
=================================

A Contao 4 extension that provides Microsoft Office template processing for Isotope eCommerce.

## Features

- Provides a new document type "Microsoft Word template" that processes a DOCX template.

## Installation

Choose the installation method that matches your workflow!


### Installation via Contao Manager

Search for `terminal42/isotope_ooxml_document` in the Contao Manager and add it to your installation. Finally,
update the packages.

### Manual installation

Add a composer dependency for this bundle. Therefore, change in the project root and run the following:

```bash
composer require terminal42/isotope_ooxml_document
```

Depending on your environment, the command can differ, i.e. starting with `php composer.phar …` if you do not have 
composer installed globally.

Then, update the database via the Contao install tool.


## Configuration

1. Place a Word file (.doc or .docx) that holds placeholder variables somewhere in the `/files` directory.
2. Create a new document of the type "Microsoft Word template" and choose the beforementioned Word file.

### Placeholder variables

Placeholders in Word files are written like: `${order_date}`. You have access to all variables as for the notifications.

You can find an [example Word template](src/Resources/docs/example.docx) in the Resources folder.

## License

This bundle is released under the [MIT license](LICENSE)
