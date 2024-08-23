# Oracle PL/SQL API for Microsoft Word DocX Documents Generation
Microsoft Word Documents API provides functionality to quickly and efficiently generate Microsoft Word Documents (DOCX) directly from Oracle database.

It requires no additional resources and it is developed in pure PL/SQL.

## Changelog
- 1.0 - Initial Release
- 2.0 - Images; table borders
- 2.1 - Default spelling and grammar language
- 2.11 - Fixed special characters issue
- 2.2 - Draw a table in the header or footer
- 2.3 - Newline in a paragraph support

## Install
- download 2 script files from "package" folder 
- execute them in database schema in following order:
1. PKS script file (package definition)
2. PKB file (package body)

New Package ZT_WORD is created in database schema.

## How to use
A script with examples is available in "examples" folder. Strongly recommended to try it first.

## Examples
Examples are wrapped up in a database procedure named p_create_word and it's source can be located in the "examples" folder (file p_create_word.sql).

Optionally there is an APEX application which You may use as a GUI. It can be also located in the examples folder (file APEX_app.sql, APEX 19.2 or newer required).

*A remark: If You want to download the file from the APEX app then uncomment the line of code "apex_application.stop_apex_engine;" in the package body (procedure p_download_document). Otherwise the download won't work.*

## Demo Application
https://apex.oracle.com/pls/apex/f?p=zttechdemo

![](https://github.com/zorantica/plsql-word/blob/master/preview.png)
