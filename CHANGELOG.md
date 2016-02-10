# Version History

## v2.0

### BC Breaks

When sending file to the user with `$ExcelWriter->outputFile($filename, $format)`: 

- `$filename` should include the extension, and spaces in the filename are **not** replaced with underscores anymore.
- `$format` should not be `xls|xlsx` anymore but a value from `\PHPExcel_IOFactory::$_autoResolveClasses`.

## v1.0 to v1.0.2

Initial version with bugfixes.