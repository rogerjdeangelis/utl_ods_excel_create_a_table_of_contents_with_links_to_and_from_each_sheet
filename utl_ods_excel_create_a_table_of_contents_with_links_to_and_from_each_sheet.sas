Ods excel create a table of contents with links to and from each sheet

see github
https://tinyurl.com/yakn2ro9
https://github.com/rogerjdeangelis/utl_ods_excel_create_a_table_of_contents_with_links_to_and_from_each_sheet

see
https://tinyurl.com/ychfo9t9
https://stackoverflow.com/questions/50930408/creating-a-link-to-the-top-of-the-page-for-each-tab-of-an-excel-spreadsheet-in-s


INPUT
=====

  sashelp.class(where=(sex="M"))
  sashelp.class(where=(sex="F"))


EXAMPLE OUTPUTS

   d:/xls/utl_ods_excel_create_a_table_of_contents_with_links_to_and_from_each_sheet.xlsx

   First sheet

      +--------------------------------------+
      |     A      |    B       |     C      |
      +--------------------------------------+
   1  |                                      |
   2  |   Detail Report of Males [link]      |
   3  |                                      |
   4  |       Data Set SASHELP.CLASS         |
   5  |                                      |
   6  |                                      |
   7  |   Detail Report of Females [link]    |
   8  |                                      |
   9  |       Data Set SASHELP.CLASS         |
  10  |                                      |
      +--------------------------------------+

  [Table of contents]

  Second Sheet

      +----------------------------------------------------------------+
      |     A      |    B       |     C      |    D       |    E       |
      +----------------------------------------------------------------+
   1  |                                                                |
   2  |  Return to Table of contents [link]                            |
   3  |                                                                |
      +----------------------------------------------------------------+
   4  | ALFRED     |    M       |    14      |    69      |  112.5     |
      +------------+------------+------------+------------+------------+
       ...
      +------------+------------+------------+------------+------------+
   24 | WILLIAM    |    M       |    15      |   66.5     |  112       |
      +------------+------------+------------+------------+------------+

  [TABLE 1]


  Third Sheet

      +----------------------------------------------------------------+
      |     A      |    B       |     C      |    D       |    E       |
      +----------------------------------------------------------------+
   1  |                                                                |
   2  |  Return to Table of contents [link]                            |
   3  |                                                                |
      +----------------------------------------------------------------+
   4  | ALICE      |    F       |    14      |    69      |  112.5     |
      +------------+------------+------------+------------+------------+
       ...
      +------------+------------+------------+------------+------------+
   24 | ANN        |    F       |    15      |   66.5     |  112       |
      +------------+------------+------------+------------+------------+

  [TABLE 2]


PROCESS
=======

%utlfkil(d:/xls/d:/xls/utl_ods_excel_create_a_table_of_contents_with_links_to_and_from_each_sheet.xlsx);

ods excel file="d:/xls/utl_ods_excel_create_a_table_of_contents_with_links_to_and_from_each_sheet.xlsx"
options(embedded_titles="yes" contents="yes");

ods proclabel= "Detail Report of Males";

proc print data=sashelp.class(where=(sex="M"));
title link="#'The Table of Contents'!a1"  "Return to TOC";
run;

ods proclabel= "Detail Report of Females";

proc print data=sashelp.class(where=(sex="F"));
title link="#'The Table of Contents'!a1"  "Return to TOC";
run;

ods excel close;

*                _               _       _
 _ __ ___   __ _| | _____     __| | __ _| |_ __ _
| '_ ` _ \ / _` | |/ / _ \   / _` |/ _` | __/ _` |
| | | | | | (_| |   <  __/  | (_| | (_| | || (_| |
|_| |_| |_|\__,_|_|\_\___|   \__,_|\__,_|\__\__,_|

;

     sashelp.class(where=(sex="M"))
     sashelp.class(where=(sex="F"))

*          _       _   _
 ___  ___ | |_   _| |_(_) ___  _ __
/ __|/ _ \| | | | | __| |/ _ \| '_ \
\__ \ (_) | | |_| | |_| | (_) | | | |
|___/\___/|_|\__,_|\__|_|\___/|_| |_|

;

* same as process;

%utlfkil(d:/xls/d:/xls/utl_ods_excel_create_a_table_of_contents_with_links_to_and_from_each_sheet.xlsx);

ods excel file="d:/xls/utl_ods_excel_create_a_table_of_contents_with_links_to_and_from_each_sheet.xlsx"
options(embedded_titles="yes" contents="yes");

ods proclabel= "Detail Report of Males";

proc print data=sashelp.class(where=(sex="M"));
title link="#'The Table of Contents'!a1"  "Return to TOC";
run;

ods proclabel= "Detail Report of Females";

proc print data=sashelp.class(where=(sex="F"));
title link="#'The Table of Contents'!a1"  "Return to TOC";
run;

ods excel close;


