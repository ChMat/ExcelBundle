# ChMatExcelBundle

This bundle integrates phpExcel library into a Symfony bundle. 
It lets you read (and write) simple Excel|CSV files.

## 📕 Archived repository

This repository is not maintained anymore.

## Installation

### Add the Bundle to your `composer.json`

Add the bundle reference in the `require` section:

    "require": {
    // ...
    "chmat/excel-bundle": "~2.0" // check Versions section for details on version numbering
    // ...
    }

### Update your Dependencies

    $ composer update

### Enable the Bundle in `AppKernel.php`

Modify your `app/AppKernel.php` file by declaring this bundle.

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

This bundle declares two services for your enjoyment. You don't have to declare them yourself.

    services:
        chmat.excel_reader:
            class: ChMat\ExcelBundle\Service\ExcelReader
        chmat.excel_writer:
            class: ChMat\ExcelBundle\Service\ExcelWriter



## Contributing

Pull requests are welcome !

## Versions

Since v2.0, version numbers are constructed with the following principles.

A valid version number will always be in the form of x.y.z. An increment of any of the value means that:

- **x** − major changes: WILL include breaking changes most of the time
- **y** − new features: MAY include breaking changes
- **z** − bugfixes and updates to the documentation: WILL NOT include breaking changes

For more information on available version numbers, please refer to the [repository on GitHub](https://github.com/ChMat/ExcelBundle/).

## License

This bundle is subject to the MIT license that is bundled with this source code in the LICENSE file.
