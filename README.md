# Component Spreadsheet
Class and method for read / generate spreadsheet files.

The supported format:
* **CSV**: parameters by default, delimiter: `;`, enclore: `"`.
    * On parsing method, convert empty value and `'NULL'` (phpmyadmin export NULL string in csv export) to `null`
* **XLSX.XML**: Xml file for Excel (excel document)


## Prerequisite

* PHP 7.4+
* PHP extension
    * Mbstring
    * Pdo mysql _(optional)_

## Install
Edit your [composer.json](https://getcomposer.org) (launch `composer update` after edit):
```json
{
  "repositories": [
    { "type": "git", "url": "git@github.com:jgauthi/component_spreadsheet.git" }
  ],
  "require": {
    "jgauthi/component_spreadsheet": "1.*"
  }
}
```


## Documentation
You can look at [folder example](https://github.com/jgauthi/component_spreadsheet/tree/master/example).

