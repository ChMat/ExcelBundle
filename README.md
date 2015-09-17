# ChMatExcelBundle

This bundle integrates phpExcel library into a Symfony bundle. 
It lets you read (and write) simple Excel|CSV files.

## Installation

### Add the Bundle to your `composer.json`

Add the bundle reference in the `require` section:

    "require": {
    // ...
    "chmat/excel-bundle": "~1.0" // check packagist.org for other tags
    // ...
    }

### Update your Dependencies

    $ composer update

### Enable the Bundle in `AppKernel.php`

    // In app/AppKernel.php
    <?php
    
    public function registerBundles()
    {
        $bundles = array(
            // ...
            new ChMat\ExcelBundle\ChMatExcelBundle(),
            // ...
        );
    }
    
## Usage

**To be updated.**

This bundle declares two services for your enjoyment:

    services:
        chmat.excel_reader:
            class: ChMat\ExcelBundle\Service\ExcelReader
        chmat.excel_writer:
            class: ChMat\ExcelBundle\Service\ExcelWriter



## Contributing

Pull requests are welcome !

## License

This bundle is subject to the MIT license that is bundled with this source code in the LICENSE file.