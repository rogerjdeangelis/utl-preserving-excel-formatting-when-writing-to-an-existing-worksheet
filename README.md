# utl-preserving-excel-formatting-when-writing-to-an-existing-worksheet
Preserving excel formatting when writing to an existing worksheet.

    Preserving excel formatting when writing to an existing worksheet.

      Problem: Update ge in Place (should work with a shared workbook)

        Change Alfreds age to 99 in worksheet 'class$n' or excel named range 'class'

      Two solutions (only works with moderately formatter sheets better if you have a primary key)

        1. Datastep modify
        2. SQL update (sas passthru - should be very fast?)

    I would stick with R's XLConnect option, STYLE_ACTION=NONE,
    when writing data to a pre-formatted Excel template.

    You may need classic SAS for this, not EG server!

    If the formatting creates valid column names and the sheet is not overly formatted, ie does not
    a piture of you mother is not between in the rows, then the following may work.

    github
    https://tinyurl.com/yxgwdvol
    https://github.com/rogerjdeangelis/utl-preserving-excel-formatting-when-writing-to-an-existing-worksheet

    github
    https://tinyurl.com/y5qq56m7
    https://github.com/rogerjdeangelis/utl-in-palce-updates-to-an-existing-shared-excel-workbook

    Other excel repos
    https://tinyurl.com/ybnm6azh
    https://github.com/rogerjdeangelis?utf8=%E2%9C%93&tab=repositories&q=excel+in%3Aname&type=&language=

    SAS Forum
    https://tinyurl.com/y4ge3h4f
    https://communities.sas.com/t5/SASware-Ballot-Ideas/ODS-EXCEL-preserve-existing-file-formats/idi-p/559408

    *_                   _
    (_)_ __  _ __  _   _| |_
    | | '_ \| '_ \| | | | __|
    | | | | | |_) | |_| | |_
    |_|_| |_| .__/ \__,_|\__|
            |_|
    ;

    * create an excel workbook with formatting;

    ods escapechar="^";
    ods excel file="d:/xls/styled.xlsx" style=statistical options(autofilter="all" frozen_headers="yes" sheet_name="class");;
    proc report data=sashelp.class style(header)={font_weight=bold font_size=14 font_style=italic} split="/";
    cols name age sex ;
    define name / style={font_style=italic font_weight=bold color=red};
    define age  /  style={ font_weight=bold color=blue};
    define sex  / display "Gender";
    run;quit;
    ods excel close;

     d:/xls/styled.xlsx

         +----------------------------+---------------------+-------------------------------+
         |          A                 |        B            |                C              |
         |----------------------------+---------------------+-------------------------------+
         || |__ | |_   _  ___         |     _ __ ___  __| | |    | |__ | | __ _  ___| | __  | Color and style preserved
         || '_ \| | | | |/ _ \        |    | '__/ _ \/ _` | |    | '_ \| |/ _` |/ __| |/ /  |
      1  || |_) | | |_| |  __/        |    | | |  __/ (_| | |    | |_) | | (_| | (__|   <   |
         ||_.__/|_|\__,_|\___|        |    |_|  \___|\__,_| |    |_.__/|_|\__,_|\___|_|\_\  |
         | _ __   __ _ _ __ ___   ___ |     __ _  __ _  ___ |       ___  _____  __          |
         || '_ \ / _` | '_ ` _ \ / _ \|    / _` |/ _` |/ _ \|      / __|/ _ \ \/ /          |
         || | | | (_| | | | | | |  __/|   | (_| | (_| |  __/|      \__ \  __/>  <           |
         ||_| |_|\__,_|_| |_| |_|\___||    \__,_|\__, |\___||      |___/\___/_/\_\          |
         |                        [V] |                  [V]|                            [V]| Auto filters preserved
         +----------------------------+---------------------+-------------------------------+
      2  |       Alfred(italic)       |        14(bold)     |              M                |
         +----------------------------+---------------------+-------------------------------+
      3  |       Alice(italic)        |        13(bold)     |              F                |
         +----------------------------+---------------------+-------------------------------+
      4  |       Barbara(italic)      |        13(bold)     |              F                |
         +----------------------------+---------------------+-------------------------------+
      5  |       Carol(italic)        |        14(bold)     |              F                |
         +----------------------------+---------------------+-------------------------------+
      6  |       Henry(italic)        |        14(bold)     |              M                |
         +----------------------------+---------------------+-------------------------------+

       [CLASS]
    *           _
     _ __ _   _| | ___  ___
    | '__| | | | |/ _ \/ __|
    | |  | |_| | |  __/\__ \
    |_|   \__,_|_|\___||___/

    ;

    Change Alfreds age to 99

    *            _               _
      ___  _   _| |_ _ __  _   _| |_
     / _ \| | | | __| '_ \| | | | __|
    | (_) | |_| | |_| |_) | |_| | |_
     \___/ \__,_|\__| .__/ \__,_|\__|
                    |_|


     d:/xls/styled.xlsx

         +----------------------------+---------------------+-------------------------------+
         |          A                 |        B            |                C              |
         |----------------------------+---------------------+-------------------------------+
         || |__ | |_   _  ___         |     _ __ ___  __| | |    | |__ | | __ _  ___| | __  | Color and style preserved
         || '_ \| | | | |/ _ \        |    | '__/ _ \/ _` | |    | '_ \| |/ _` |/ __| |/ /  |
      1  || |_) | | |_| |  __/        |    | | |  __/ (_| | |    | |_) | | (_| | (__|   <   |
         ||_.__/|_|\__,_|\___|        |    |_|  \___|\__,_| |    |_.__/|_|\__,_|\___|_|\_\  |
         | _ __   __ _ _ __ ___   ___ |     __ _  __ _  ___ |       ___  _____  __          |
         || '_ \ / _` | '_ ` _ \ / _ \|    / _` |/ _` |/ _ \|      / __|/ _ \ \/ /          |
         || | | | (_| | | | | | |  __/|   | (_| | (_| |  __/|      \__ \  __/>  <           |
         ||_| |_|\__,_|_| |_| |_|\___||    \__,_|\__, |\___||      |___/\___/_/\_\          |
         |                        [V] |                  [V]|                            [V]| Auto filters preserved
         +----------------------------+---------------------+-------------------------------+
      2  |       Alfred(italic)       |        99(bold)     |              M                | Alfreds age changed to 99
         +----------------------------+---------------------+-------------------------------+ color and style preserved
      3  |       Alice(italic)        |        13(bold)     |              F                |
         +----------------------------+---------------------+-------------------------------+
      4  |       Barbara(italic)      |        13(bold)     |              F                |
         +----------------------------+---------------------+-------------------------------+
      5  |       Carol(italic)        |        14(bold)     |              F                |
         +----------------------------+---------------------+-------------------------------+
      6  |       Henry(italic)        |        14(bold)     |              M                |
         +----------------------------+---------------------+-------------------------------+

       [CLASS]
     d:/xls/styled.xlsx (updated in place)

         +----------------------------+---------------------+-------------------------------+
         |          A                 |        B            |                C              |
         |----------------------------+---------------------+-------------------------------+
         || |__ | |_   _  ___         |     _ __ ___  __| | |    | |__ | | __ _  ___| | __  |
         || '_ \| | | | |/ _ \        |    | '__/ _ \/ _` | |    | '_ \| |/ _` |/ __| |/ /  |
      1  || |_) | | |_| |  __/        |    | | |  __/ (_| | |    | |_) | | (_| | (__|   <   |
         ||_.__/|_|\__,_|\___|        |    |_|  \___|\__,_| |    |_.__/|_|\__,_|\___|_|\_\  |
         | _ __   __ _ _ __ ___   ___ |     __ _  __ _  ___ |       ___  _____  __          |
         || '_ \ / _` | '_ ` _ \ / _ \|    / _` |/ _` |/ _ \|      / __|/ _ \ \/ /          |
         || | | | (_| | | | | | |  __/|   | (_| | (_| |  __/|      \__ \  __/>  <           |
         ||_| |_|\__,_|_| |_| |_|\___||    \__,_|\__, |\___||      |___/\___/_/\_\          |
         +----------------------------+---------------------+-------------------------------+
      2  |       Alfred(italic)       |        14(bold)     |              M                |
         +----------------------------+---------------------+-------------------------------+
      3  |       Alice(italic)        |        13(bold)     |              F                |
         +----------------------------+---------------------+-------------------------------+
      4  |       Barbara(italic)      |        13(bold)     |              F                |
         +----------------------------+---------------------+-------------------------------+
      5  |       Carol(italic)        |        14(bold)     |              F                |
         +----------------------------+---------------------+-------------------------------+
      6  |       Henry(italic)        |        14(bold)     |              M                |
         +----------------------------+---------------------+-------------------------------+

       [CLASS]


     *          _       _   _
     ___  ___ | |_   _| |_(_) ___  _ __  ___
    / __|/ _ \| | | | | __| |/ _ \| '_ \/ __|
    \__ \ (_) | | |_| | |_| | (_) | | | \__ \
    |___/\___/|_|\__,_|\__|_|\___/|_| |_|___/

    ;

    ********************
    1. Datastep modify *
    ********************

    libname xel "d:/xls/styled.xlsx" scan_text=no;
       data Xel.'class$'n ;
          modify Xel.'class$'n ;;
          age=99;
          where name= 'Alfred';
       run;quit;
    libname xel clear;

    NOTE: The data set XEL.class has been updated.
    There were 1 observations rewritten, 0 observations added and 0 observations deleted.


    ********************************************************
    2. SQL update (sas passthru - should be very fast?)    *
    ********************************************************

     SAS/Passthru

    If you make a named range 'class' then this will work

    * this does it in place;
    proc sql dquote=ansi;
       connect to excel as excel(Path="d:/xls/styled.xlsx");
       execute(
         update Xel.'class$'n ;
         set age=88
         where name="Alfred"
       ) by excel;
       disconnect from excel;
    Quit;

