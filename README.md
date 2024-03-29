# Oracle PL/SQL API for Microsoft Word DocX Documents Generation
Microsoft Word Documents Generator Package provides functionality to quickly and efficiently generate Microsoft Word Documents (DOCX) directly from Oracle database.

It requires no additional resources and it is developed in pure PL/SQL.

## Changelog
1.0 - Initial Release

2.0 - Images; table borders

2.1 - Default spelling and grammar language

2.11 - Fixed special characters issue

2.2 - Draw a table in the header or footer

## Install
- download 2 script files from "package" directory 
- execute them in database schema in following order:
1. PKS script file (package definition)
2. PKB file (package body)

New Package ZT_WORD is created in database schema.

## How to use
A script with examples is available in "examples" directory. Strongly recommended to try it first.

## Demo Application
https://apex.oracle.com/pls/apex/f?p=zttechdemo

![](https://github.com/zorantica/plsql-word/blob/master/preview.png)
