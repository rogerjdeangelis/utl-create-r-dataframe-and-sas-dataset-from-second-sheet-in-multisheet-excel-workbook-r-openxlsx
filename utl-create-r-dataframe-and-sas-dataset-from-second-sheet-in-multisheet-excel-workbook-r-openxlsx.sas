%let pgm=utl-create-r-dataframe-and-sas-dataset-from-second-sheet-in-multisheet-excel-workbook-r-openxlsx;

%stop_submission;

Create r dataframe and sas dataset from second sheet in multisheet excel workbook r openxlsx

  CONTENTS

     1 import second sheet
     2 Related repos on end

github
https://tinyurl.com/mt5y4bpw
https://github.com/rogerjdeangelis/utl-create-r-dataframe-and-sas-dataset-from-second-sheet-in-multisheet-excel-workbook-r-openxlsx

communities.sas
https://tinyurl.com/yehj3xe6
https://communities.sas.com/t5/SAS-Programming/always-import-second-sheet-from-multiple-excel-files-without/m-p/826649#M326514


/**************************************************************************************************************************/
/*             INPUT                       |      PROCESS                         |      OUTPUT                           */
/*             =====                       |      =======                         |      ======                           */
/*                                         |                                      |                                       */
/* d:/xls/sheets.xlsx                      | 1 IMPORT SECOND SHEET                | SAS                                   */
/*                                         |                                      |                                       */
/* SECOND SHEET                            | want <- read.xlsx(                   | SD1.WANT (IMPORTED 2ND SHEET)         */
/* ============                            |   "d:/xls/sheets.xlsx"               |                                       */
/*                                         |   ,sheet = 2)                        | OBS    NAME       SEX    AGE          */
/* -----------------------+                |                                      |                                       */
/* | A1| fx    |NAME      |                | You can aslo getall sheets           |   1    Alfred      M      14          */
/* -----------------------------------+    |                                      |   5    Henry       M      14          */
/* [_] |    A     |    B    |    C    |    | sheet_names <- getSheetNames(file)   |   6    James       M      12          */
/* -----------------------------------|    |                                      |   9    Jeffrey     M      13          */
/*  1  | NAME     |   SEX   |   AGE   |    |                                      |  10    John        M      12          */
/*  -- |----------+---------+---------|    | %utl_rbeginx;                        |                                       */
/*  2  |Alfred    | M       | 14      |    | parmcards4;                          |  R                                    */
/*  -- |----------+---------+---------|    | library(haven)                       |    Obs    NAME SEX AGE                */
/*  3  |Henry     | M       | 14      |    | library(openxlsx)                    |  1   1  Alfred   M  14                */
/*  -- |----------+---------+---------|    | source("c:/oto/fn_tosas9x.R")        |  2   5   Henry   M  14                */
/*  4  |James     | M       | 12      |    | want <- read.xlsx(                   |  3   6   James   M  12                */
/*  -- |----------+---------+---------|    |   "d:/xls/sheets.xlsx"               |  4   9 Jeffrey   M  13                */
/*  5  |Jeffrey   | M       | 13      |    |   ,sheet = 2)                        |  5  10    John   M  12                */
/*  -- |----------+---------+---------|    | want                                 |                                       */
/*  6  |John      | M       | 12      |    | fn_tosas9x(                          |                                       */
/*  -- |----------+---------+---------|    |       inp    = want                  |                                       */
/* [MALES]                                 |      ,outlib ="d:/sd1/"              |                                       */
/*                                         |      ,outdsn ="want"                 |                                       */
/* FIRST SHEET                             |      )                               |                                       */
/* ============                            | ;;;;                                 |                                       */
/*                                         | %utl_rendx;                          |                                       */
/* -----------------------+                |                                      |                                       */
/* | A1| fx    |DAYNUM    |                | proc print data=sd1.want;            |                                       */
/* -----------------------------------+    | run;quit;                            |                                       */
/* [_] |    A     |    B    |    C    |    |                                      |                                       */
/* -----------------------------------|    |                                      |                                       */
/*  1  | NAME     |   SEX   |   AGE   |    |                                      |                                       */
/*  -- |----------+---------+---------|    |                                      |                                       */
/*  2  |Alice     | F       | 13      |    |                                      |                                       */
/*  -- |----------+---------+---------|    |                                      |                                       */
/*  3  |Barbara   | F       | 13      |    |                                      |                                       */
/*  -- |----------+---------+---------|    |                                      |                                       */
/*  4  |Carol     | F       | 14      |    |                                      |                                       */
/*  -- |----------+---------+---------|    |                                      |                                       */
/*  5  |Jane      | F       | 12      |    |                                      |                                       */
/*  -- |----------+---------+---------|    |                                      |                                       */
/*  6  |Janet     | F       | 15      |    |                                      |                                       */
/*  -- |----------+---------+---------|    |                                      |                                       */
/* [FEMALES]                               |                                      |                                       */
/*                                         |                                      |                                       */
/*                                         |                                      |                                       */
/* ods excel file="d:/xls/sheets.xlsx" ;   |                                      |                                       */
/* ods excel                               |                                      |                                       */
/*   options(sheet_name="Females");        |                                      |                                       */
/* proc print                              |                                      |                                       */
/*    data=sashelp.class(                  |                                      |                                       */
/*      obs=5                              |                                      |                                       */
/*      keep=name age sex                  |                                      |                                       */
/*      where=( sex='F'));                 |                                      |                                       */
/* run;quit;                               |                                      |                                       */
/* ods excel options(                      |                                      |                                       */
/*    sheet_name="Males"                   |                                      |                                       */
/*    sheet_interval="NOW");               |                                      |                                       */
/* proc print                              |                                      |                                       */
/*    data=sashelp.class(                  |                                      |                                       */
/*      obs=5                              |                                      |                                       */
/*      keep=name age sex                  |                                      |                                       */
/*      where=( sex='M')) ;                |                                      |                                       */
/* run;quit;;                              |                                      |                                       */
/* ods excel close;                        |                                      |                                       */
/* run;quit;                               |                                      |                                       */
/**************************************************************************************************************************/

/*                   _
(_)_ __  _ __  _   _| |_
| | `_ \| `_ \| | | | __|
| | | | | |_) | |_| | |_
|_|_| |_| .__/ \__,_|\__|
        |_|
*/

 ods excel file="d:/xls/sheets.xlsx" ;
 ods excel
   options(sheet_name="Females");
 proc print
    data=sashelp.class(
      obs=5
      keep=name age sex
      where=( sex='F'));
 run;quit;
 ods excel options(
    sheet_name="Males"
    sheet_interval="NOW");
 proc print
    data=sashelp.class(
      obs=5
      keep=name age sex
      where=( sex='M')) ;
 run;quit;;
 ods excel close;
 run;quit;


 d:/xls/sheets.xlsx

 FIRST SHEET
 ==========

 -----------------------+
 | A1| fx    |NAME      |
 -----------------------------------+
 [_] |    A     |    B    |    C    |
 -----------------------------------|
  1  | NAME     |   SEX   |   AGE   |
  -- |----------+---------+---------|
  2  |Alfred    | M       | 14      |
  -- |----------+---------+---------|
  3  |Henry     | M       | 14      |
  -- |----------+---------+---------|
  4  |James     | M       | 12      |
  -- |----------+---------+---------|
  5  |Jeffrey   | M       | 13      |
  -- |----------+---------+---------|
  6  |John      | M       | 12      |
  -- |----------+---------+---------|
 [MALES]

 SECOND SHEET
 ============

 -----------------------+
 | A1| fx    |DAYNUM    |
 -----------------------------------+
 [_] |    A     |    B    |    C    |
 -----------------------------------|
  1  | NAME     |   SEX   |   AGE   |
  -- |----------+---------+---------|
  2  |Alice     | F       | 13      |
  -- |----------+---------+---------|
  3  |Barbara   | F       | 13      |
  -- |----------+---------+---------|
  4  |Carol     | F       | 14      |
  -- |----------+---------+---------|
  5  |Jane      | F       | 12      |
  -- |----------+---------+---------|
  6  |Janet     | F       | 15      |
  -- |----------+---------+---------|
 [FEMALES]

/*
 _ __  _ __ ___   ___ ___  ___ ___
| `_ \| `__/ _ \ / __/ _ \/ __/ __|
| |_) | | | (_) | (_|  __/\__ \__ \
| .__/|_|  \___/ \___\___||___/___/
|_|
*/

%utl_rbeginx;
parmcards4;
library(haven)
library(openxlsx)
source("c:/oto/fn_tosas9x.R")
want <- read.xlsx(
  "d:/xls/sheets.xlsx"
  ,sheet = 2)
want
fn_tosas9x(
      inp    = want
     ,outlib ="d:/sd1/"
     ,outdsn ="want"
     )
;;;;
%utl_rendx;

proc print data=sd1.want;
run;quit;

/**************************************************************************************************************************/
/*      OUTPUT                                                                                                            */
/*      ======                                                                                                            */
/*                                                                                                                        */
/* SAS                                                                                                                    */
/*                                                                                                                        */
/* SD1.WANT (IMPORTED 2ND SHEET)                                                                                          */
/*                                                                                                                        */
/* OBS    NAME       SEX    AGE                                                                                           */
/*                                                                                                                        */
/*   1    Alfred      M      14                                                                                           */
/*   5    Henry       M      14                                                                                           */
/*   6    James       M      12                                                                                           */
/*   9    Jeffrey     M      13                                                                                           */
/*  10    John        M      12                                                                                           */
/*                                                                                                                        */
/*  R                                                                                                                     */
/*    Obs    NAME SEX AGE                                                                                                 */
/*  1   1  Alfred   M  14                                                                                                 */
/*  2   5   Henry   M  14                                                                                                 */
/*  3   6   James   M  12                                                                                                 */
/*  4   9 Jeffrey   M  13                                                                                                 */
/*  5  10    John   M  12                                                                                                 */
/**************************************************************************************************************************/


https://github.com/rogerjdeangelis/Export-exel-sheet-names-to-sas-dataset
https://github.com/rogerjdeangelis/utl-Import-excel-sheet-as-character-fixing-truncation-mixed-type-columns-and-appending-issues
https://github.com/rogerjdeangelis/utl-add-monthly-worksheets-to-an-existing-yearly-excel-workbook
https://github.com/rogerjdeangelis/utl-adding-a-second-ods-excel-created-sheet-to-a-closed-ods-excel-workbook
https://github.com/rogerjdeangelis/utl-adding-a-sheet-to-an-existing-open-and-on-screen-or-saved-excel-workbook
https://github.com/rogerjdeangelis/utl-appending-records-to-an-existing-excel-sheet
https://github.com/rogerjdeangelis/utl-apply-excel-styling-across-multiple-spreadsheets-using-openxlsx-in-r
https://github.com/rogerjdeangelis/utl-creating-multiple-odbc-tables-in-a-one-excel-sheet
https://github.com/rogerjdeangelis/utl-does-the-excel-sheet-exist
https://github.com/rogerjdeangelis/utl-excel-grid-of-four-reports-in-one-sheet
https://github.com/rogerjdeangelis/utl-excel-hiding-columns-height-and-weight-in-sheet-class
https://github.com/rogerjdeangelis/utl-excel-use-the-name-of-the-last-variable-in-the-pdv-for-sheet-name
https://github.com/rogerjdeangelis/utl-extract-sheet-names-from-multiple-excel-versions-using-r
https://github.com/rogerjdeangelis/utl-extracting-hyperlinks-from-an-excel-sheet-python
https://github.com/rogerjdeangelis/utl-highlight-existing-cells-in-excel-sheet2-that-correspond-to-cells-in-sheet1-with-specified-value
https://github.com/rogerjdeangelis/utl-how-to-check-whether-a-student-is-in-the-Excel-sheet-class
https://github.com/rogerjdeangelis/utl-how-to-load-excel-sheets-with-sheet-names-with-31-characters
https://github.com/rogerjdeangelis/utl-importing-excel-when-sheetname-has-spaces
https://github.com/rogerjdeangelis/utl-importing-multiple-excel-worksheets-without-access-to-pc-files
https://github.com/rogerjdeangelis/utl-ods-excel-update-excel-sheet-in-place-python
https://github.com/rogerjdeangelis/utl-pivot-long--excel-sheet-and-run-a-regression-in-r-and-python
https://github.com/rogerjdeangelis/utl-pivot-transpose-an-excel-sheet-with-columns-that-are-excel-dates
https://github.com/rogerjdeangelis/utl-preserving-excel-formatting-when-writing-to-an-existing-worksheet
https://github.com/rogerjdeangelis/utl-programatically-search-all-cells-in-an-excel-sheet-for-an-arbitrary-string-python-openxl
https://github.com/rogerjdeangelis/utl-remove-sheet-from-excel-workbook
https://github.com/rogerjdeangelis/utl-remove-sheet-from-existing-excel-worksheet-unix-and-windows-R
https://github.com/rogerjdeangelis/utl-sending-a-formula-to-excel-to-reference-a-cell-in-another-sheet
https://github.com/rogerjdeangelis/utl-side-by-side-reports-within-arbitrary-positions-in-one-excel-sheet-wps-r
https://github.com/rogerjdeangelis/utl-side-by-side-sas-tables-in-one-excel-sheet
https://github.com/rogerjdeangelis/utl-update-a-master-sheet-with-transaction-sheet-using-excel-and-r-openxls-package-and-sqldf
https://github.com/rogerjdeangelis/utl-update-existing-excel-sheet-in-place-using-r-dcom-client
https://github.com/rogerjdeangelis/utl-update-in-place-sheet2-by-adding-dinner-costs-from-sheet1-preserving-excel-formatting-r
https://github.com/rogerjdeangelis/utl-using-only-r-openxlsx-to-add-excel-formulas-to-an-existing-sheet
https://github.com/rogerjdeangelis/utl-very-fast-summation-of-ll6-columns-in-excel-without-importing-to-sheet
https://github.com/rogerjdeangelis/utl_adding_SAS_graphics_at_an_arbitrary_position_into_existing_excel_sheets
https://github.com/rogerjdeangelis/utl_appending_workbook2_sheet1_to_worlbook1_sheet1
https://github.com/rogerjdeangelis/utl_excel_add_sheet
https://github.com/rogerjdeangelis/utl_excel_add_to_sheet
https://github.com/rogerjdeangelis/utl_excel_combining_sheets_without_common_names_types_lengths
https://github.com/rogerjdeangelis/utl_excel_create_a_sheet_for_each_table_with_variable_name_position_and_label
https://github.com/rogerjdeangelis/utl_excel_import_two_excel_ranges_within_one_sheet
https://github.com/rogerjdeangelis/utl_excel_merge_two-sheets
https://github.com/rogerjdeangelis/utl_excel_using_byval_sex_and_sheet_interval_bygroups_to_create_multiple_worksheets
https://github.com/rogerjdeangelis/utl_import_data_from_excel_sheet_with_headers_and_footers_without_specifying_range_option
https://github.com/rogerjdeangelis/utl_importing_multiple_pwd_protected_xlm_workbook_and_sheets
https://github.com/rogerjdeangelis/utl_importing_three_excel_tables_that_are_in_one_sheet
https://github.com/rogerjdeangelis/utl_joining_and_updating_excel_sheets_without_importing_data
https://github.com/rogerjdeangelis/utl_link_student_names_in_index_worksheet_to_detail_data_in_class_worksheet
https://github.com/rogerjdeangelis/utl_maintaining_all_significant_digits_when_importing_excel_sheet
https://github.com/rogerjdeangelis/utl_ods_excel_create_a_table_of_contents_with_links_to_and_from_each_sheet
https://github.com/rogerjdeangelis/utl_put_excel_sheetnames_into_sas_macro_variable
https://github.com/rogerjdeangelis/utl_table_of_contents_with_excel_links_to_sheets


/*              _
  ___ _ __   __| |
 / _ \ `_ \ / _` |
|  __/ | | | (_| |
 \___ _| |_|\__,_|

*/
