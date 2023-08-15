# Component Spreadsheet
Executables, class and method for read / generate spreadsheet files.

The supported format:
* **CSV**: parameters by default, delimiter: `;`, enclore: `"`.
    * On parsing method, convert empty value and `'NULL'` (phpmyadmin export NULL string in csv export) to `null`
* **XLSX.XML**: Xml file for Excel (excel document)


## Prerequisite
* PHP 8.2 (v1.1+) or 7.4+ (v1.0)
* PHP extension
    * Iconv
    * Mbstring
    * _(optional)_ Pdo, pdo-mysqli or another database (sqlite, etc)

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
You can look at [folder example](example).
