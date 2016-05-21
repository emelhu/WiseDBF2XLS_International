#WiseDBF2XLS_International
==========================

Convert DBF to XLS with localization.
-------------------------------------

DBF file can store contained text not only with english characters.
After the head block DBF is a simple DOS/WIN (text)file with selected charset or charset of operating system. 

Most conversion utilities can't manage charsets and demage DBF content when convert it to XLS file

This utility is a simple exe (although it use .net environment) and directly read DBF file and write XLS file.

...and more... dBaseIII is a well-known format, but I had couldn't read DBF version 4,5,6,7 with Excel or LibraOffice... this utility read these correctly.

used components:
https://github.com/emelhu/NDbfReaderEx  for read DBF & DBT
https://code.google.com/archive/p/excellibrary/  for write XLS

so the formats and limitations same as in this packages.

It's a command line utility and requires the .Net Framework 4.5.2.

Freeware by eMeL ( www.emel.hu )  MIT licence
