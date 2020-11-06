CREATE OR REPLACE PACKAGE BODY ASP_PAGE_GEN_PKG
AS
TYPE cur_typ IS REF CURSOR;
PROCEDURE create_ASP_maint_page_prc(v_table_name IN VARCHAR2)
IS
v_tbl_pfx VARCHAR2(30);
v_pkg_name VARCHAR2(30);
c cur_typ;
sqlStr VARCHAR2(1000);
validateFunctionCalls VARCHAR2(1000);
TYPE column_name_arr IS VARRAY(200) OF all_tab_columns.column_name%TYPE;
TYPE pk_constraint_type_arr IS VARRAY(200) OF all_constraints.constraint_type%TYPE;
TYPE column_type_arr IS VARRAY(200) OF all_tab_columns.data_type%TYPE;
TYPE constraint_table_arr IS VARRAY(200) OF all_tab_columns.table_name%TYPE;
TYPE nullable_arr IS VARRAY(200) OF all_tab_columns.nullable%TYPE;
TYPE pk_search_condition_arr IS VARRAY(200) OF all_constraints.search_condition%TYPE;
v_column_name_arr column_name_arr := column_name_arr();
v_column_type_arr column_type_arr := column_type_arr();
v_pk_constraint_type_arr pk_constraint_type_arr := pk_constraint_type_arr();
v_constraint_table_arr constraint_table_arr := constraint_table_arr();
v_nullable_arr nullable_arr := nullable_arr();
v_pk_search_condition_arr pk_search_condition_arr := pk_search_condition_arr();
fileHandler UTL_FILE.FILE_TYPE;
vDateFormatStr VARCHAR2(100):= '</B>';
j NUMBER :=0;
whereClause VARCHAR2(1000);
logical_delete_found NUMBER:= 0;
selectColStr VARCHAR2(1000) := NULL;
v_prev_column_name all_tab_columns.column_name%TYPE := 'XXXX';
CURSOR all_fields_cur(v_tbl_name IN VARCHAR2)
IS
SELECT c.column_name, c.data_type,pk.constraint_type, DECODE(pk.constraint_type, 'R', (SELECT DISTINCT table_name
       FROM all_constraints
       WHERE constraint_name = pk.r_constraint_name),
       NULL) constraint_table, nullable, pk.search_condition
FROM   all_tab_columns c,
(     SELECT a.table_name, b.column_name,a.constraint_type, a.r_constraint_name, a.search_condition
     FROM   all_constraints a,
            all_cons_columns  b
     WHERE a.constraint_name = b.constraint_name )pk
WHERE  c.table_name = pk.table_name(+)
AND    c.column_name = pk.column_name(+)
AND    c.table_name = v_tbl_name
ORDER  BY c.column_id ASC;

BEGIN
     v_column_name_arr.EXTEND;
     v_column_type_arr.EXTEND;
     v_pk_constraint_type_arr.EXTEND;
     v_constraint_table_arr.EXTEND;
     v_nullable_arr.EXTEND;
     v_pk_search_condition_arr.EXTEND;

     SELECT REPLACE(UPPER(v_table_name), '_TBL')||'_PKG',REPLACE(lower(v_table_name), '_tbl')
     INTO v_pkg_name, v_tbl_pfx
     FROM DUAL;
     fileHandler := UTL_FILE.FOPEN('ASP_DIR', UPPER(v_tbl_pfx) || '.asp','W');
     -- GET All columns
     OPEN all_fields_cur(v_table_name);
     FETCH all_fields_cur BULK COLLECT INTO v_column_name_arr ,
                                            v_column_type_arr ,
                                            v_pk_constraint_type_arr,
                                            v_constraint_table_arr,
                                            v_nullable_arr,
                                            v_pk_search_condition_arr;
     UTL_FILE.PUT_LINE(fileHandler,'<%@ Language=VBScript %>');
     UTL_FILE.PUT_LINE(fileHandler,'<%');
     UTL_FILE.PUT_LINE(fileHandler,'dim i, offset, allrecords, listOfColumns, LOC');
     UTL_FILE.PUT_LINE(fileHandler,'i = 0');
     UTL_FILE.PUT_LINE(fileHandler,'connstr = "DSN=ODBCTESTDB;UID=SCOTT;PWD=TIGER"');
     UTL_FILE.PUT_LINE(fileHandler,'Set oADOConnection = Server.CreateObject("ADODB.Connection")');
     UTL_FILE.PUT_LINE(fileHandler,'oADOConnection.Provider = "MSDASQL"');
     UTL_FILE.PUT_LINE(fileHandler,'oADOConnection.ConnectionString = connstr');

     UTL_FILE.PUT_LINE(fileHandler,'oADOConnection.CommandTimeout = 0');
     UTL_FILE.PUT_LINE(fileHandler,'oADOConnection.Open');
     UTL_FILE.PUT_LINE(fileHandler,'UpdType = request.querystring("Update")');
     UTL_FILE.PUT_LINE(fileHandler,'DelType = request.querystring("Delete")');
     UTL_FILE.PUT_LINE(fileHandler,'InsType = request.querystring("Insert")');
     UTL_FILE.PUT_LINE(fileHandler,'if ( request.querystring("offset") = "" ) then');
     UTL_FILE.PUT_LINE(fileHandler,'     offset = 0');
     UTL_FILE.PUT_LINE(fileHandler,'else');
     UTL_FILE.PUT_LINE(fileHandler,'     offset = request.querystring("offset")');
     UTL_FILE.PUT_LINE(fileHandler,'end if');

     UTL_FILE.PUT_LINE(fileHandler,'if UpdType = "TRUE" then');

     FOR i IN 1..v_column_name_arr.LAST
     LOOP
         IF v_prev_column_name <> v_column_name_arr(i) THEN
            UTL_FILE.PUT_LINE(fileHandler,'session("' || REPLACE(INITCAP(v_column_name_arr(i)), '_') || 'Val") = request.form("' || v_column_name_arr(i)||'")');
         END IF;
         v_prev_column_name := v_column_name_arr(i);
     END LOOP;
     UTL_FILE.PUT_LINE(fileHandler,'updSqlStr = " UPDATE ' || v_table_name || 'SET "');
     FOR i IN 1..v_column_name_arr.LAST
     LOOP
         IF v_prev_column_name <> v_column_name_arr(i) THEN
            IF v_pk_constraint_type_arr(i) = 'P' THEN
               whereClause := ' updSqlStr = updSqlStr + "WHERE ' || v_column_name_arr(i) || ' = ''" + session("' || REPLACE(INITCAP(v_column_name_arr(i)), '_')|| 'Val") +"''"';
            ELSE
                IF i = v_column_name_arr.LAST THEN
                   IF v_column_type_arr(i) = 'DATE' THEN
                       UTL_FILE.PUT_LINE(fileHandler,'updSqlStr = updSqlStr + "' || v_column_name_arr(i) || ' = TO_DATE(''" + session("' || REPLACE(INITCAP(v_column_name_arr(i)), '_') || 'Val") +"'', ''MM/DD/YYYY'')  "');
                   ELSE
                       UTL_FILE.PUT_LINE(fileHandler,'updSqlStr = updSqlStr + "' || v_column_name_arr(i) || ' = ''" + session("' || REPLACE(INITCAP(v_column_name_arr(i)), '_') || 'Val") +"''  "');
                   END IF;
                ELSE
                   IF v_column_type_arr(i) = 'DATE' THEN
                      UTL_FILE.PUT_LINE(fileHandler,'updSqlStr = updSqlStr + "'|| v_column_name_arr(i) || ' = TO_DATE(''" + session("' || REPLACE(INITCAP(v_column_name_arr(i)), '_') || 'Val") +"'', ''MM/DD/YYYY'') , "');
                   ELSE
                      UTL_FILE.PUT_LINE(fileHandler,'updSqlStr = updSqlStr + "'|| v_column_name_arr(i) || ' = ''" + session("' || REPLACE(INITCAP(v_column_name_arr(i)), '_') || 'Val") +"'' , "');
                   END IF;
                END IF;
            END IF;
         END IF;
         v_prev_column_name := v_column_name_arr(i);
     END LOOP;
     UTL_FILE.PUT_LINE(fileHandler,whereClause);
     UTL_FILE.PUT_LINE(fileHandler,'offset = request.querystring("offset")');
     UTL_FILE.PUT_LINE(fileHandler,'session("Update") = "YES"');
     UTL_FILE.PUT_LINE(fileHandler,'oADOConnection.Execute(updSqlStr)');
     UTL_FILE.PUT_LINE(fileHandler,'end if');
     UTL_FILE.PUT_LINE(fileHandler,'if DelType = "TRUE" then');
     FOR i IN 1..v_column_name_arr.LAST
     LOOP
         IF v_prev_column_name <> v_column_name_arr(i) THEN
            UTL_FILE.PUT_LINE(fileHandler,'session("' || REPLACE(INITCAP(v_column_name_arr(i)), '_') || 'Val") = request.form("' || v_column_name_arr(i)||'")');
         END IF;
         v_prev_column_name := v_column_name_arr(i);

     END LOOP;
     UTL_FILE.PUT_LINE(fileHandler,'delSqlStr = " DELETE ' || v_table_name || '"');
     FOR i IN 1..v_column_name_arr.LAST
     LOOP
         IF v_prev_column_name <> v_column_name_arr(i) THEN

            IF v_pk_constraint_type_arr(i) = 'P' THEN
               UTL_FILE.PUT_LINE(fileHandler,' delSqlStr = delSqlStr + " WHERE ' || v_column_name_arr(i) || ' = ''" + session("' || REPLACE(INITCAP(v_column_name_arr(i)), '_') || 'Val") +"''"');
            END IF;
         END IF;
         v_prev_column_name := v_column_name_arr(i);

     END LOOP;

     UTL_FILE.PUT_LINE(fileHandler,'offset = request.querystring("offset")');
     UTL_FILE.PUT_LINE(fileHandler,'session("Delete") = "YES"');
     UTL_FILE.PUT_LINE(fileHandler,'oADOConnection.Execute(delSqlStr)');
     UTL_FILE.PUT_LINE(fileHandler,'end if');
     UTL_FILE.PUT_LINE(fileHandler,'if InsType = "TRUE" then');

     FOR i IN 1..v_column_name_arr.LAST
     LOOP
         IF v_prev_column_name <> v_column_name_arr(i) THEN

            UTL_FILE.PUT_LINE(fileHandler,'session("' || REPLACE(INITCAP(v_column_name_arr(i)), '_') || 'Val") = request.form("' || v_column_name_arr(i)||'")');

            IF v_pk_constraint_type_arr(i) = 'P' THEN
               UTL_FILE.PUT_LINE(fileHandler,' getNextSequenceSql = " SELECT ' || v_column_name_arr(i) || '_SEQUENCE.NEXTVAL seq FROM DUAL "');
               UTL_FILE.PUT_LINE(fileHandler,'Set objRs = oADOConnection.Execute(getNextSequenceSql) ');
               UTL_FILE.PUT_LINE(fileHandler, v_column_name_arr(i) || 'Seq = objRs.Fields("seq").value ');
            END IF;
         END IF;
         v_prev_column_name := v_column_name_arr(i);

     END LOOP;


     UTL_FILE.PUT_LINE(fileHandler,'insSqlStr = " INSERT INTO ' || v_table_name|| ' VALUES ( "');
     FOR i IN 1..v_column_name_arr.LAST
     LOOP

         IF v_prev_column_name <> v_column_name_arr(i) THEN
            IF v_pk_constraint_type_arr(i) = 'P' THEN
               UTL_FILE.PUT_LINE(fileHandler,'insSqlStr = insSqlStr + cStr(' || v_column_name_arr(i) || 'Seq) +" , "');
            ELSE
               IF i = v_column_name_arr.LAST THEN
                  IF v_column_type_arr(i) = 'DATE' THEN
                     UTL_FILE.PUT_LINE(fileHandler,'insSqlStr = insSqlStr + " TO_DATE(''" + session("' || REPLACE(INITCAP(v_column_name_arr(i)), '_') || 'Val")+"'', ''MM/DD/YYYY'') ) "');
                  ELSE
                     UTL_FILE.PUT_LINE(fileHandler,'insSqlStr = insSqlStr + " ''" + session("' || REPLACE(INITCAP(v_column_name_arr(i)), '_') || 'Val") +"'' ) "');
                  END IF;
               ELSE
                  IF v_column_type_arr(i) = 'DATE' THEN
                     UTL_FILE.PUT_LINE(fileHandler,'insSqlStr = insSqlStr + " TO_DATE(''" + session("' || REPLACE(INITCAP(v_column_name_arr(i)), '_') || 'Val")+"'', ''MM/DD/YYYY'') , "');
                  ELSE
                     UTL_FILE.PUT_LINE(fileHandler,'insSqlStr = insSqlStr + " ''" + session("' || REPLACE(INITCAP(v_column_name_arr(i)), '_') || 'Val") +"'' , "');
                  END IF;
               END IF;
            END IF;
         END IF;
         v_prev_column_name := v_column_name_arr(i);

     END LOOP;
     UTL_FILE.PUT_LINE(fileHandler,'offset = request.querystring("offset")');
     UTL_FILE.PUT_LINE(fileHandler,'session("Insert") = "YES"');
     UTL_FILE.PUT_LINE(fileHandler,'oADOConnection.Execute(insSqlStr)');
     UTL_FILE.PUT_LINE(fileHandler,'end if');

     UTL_FILE.PUT_LINE(fileHandler,'ActionURL = "'|| v_tbl_pfx || '.asp?offset=0"');

     UTL_FILE.PUT_LINE(fileHandler,'nPage = request.querystring("Page")');

     UTL_FILE.PUT_LINE(fileHandler,'%>');
     UTL_FILE.PUT_LINE(fileHandler,'<HTML>');
     UTL_FILE.PUT_LINE(fileHandler,'<meta http-equiv="Content-Type"');
     UTL_FILE.PUT_LINE(fileHandler,'content="text/html; charset=iso-8859-1">');
     UTL_FILE.PUT_LINE(fileHandler,'<meta name="GENERATOR" content="Microsoft FrontPage Express 2.0">');
     UTL_FILE.PUT_LINE(fileHandler,'<TITLE>'||v_table_name||' MAINTENANCE</TITLE>');
     UTL_FILE.PUT_LINE(fileHandler,'<HEAD>');
     UTL_FILE.PUT_LINE(fileHandler,'<H1 ALIGN="CENTER"><B>'||v_table_name||' MAINTENANCE</B></H1>');
     UTL_FILE.PUT_LINE(fileHandler,'<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">');
     UTL_FILE.PUT_LINE(fileHandler,'<link href="style.css" rel="stylesheet" type="text/css">');
     UTL_FILE.PUT_LINE(fileHandler,'<style type="text/css">');
     UTL_FILE.PUT_LINE(fileHandler,'</style>');
     UTL_FILE.PUT_LINE(fileHandler,'<script>');
     UTL_FILE.PUT_LINE(fileHandler,'function validate_list(field, value, params) {');
     UTL_FILE.PUT_LINE(fileHandler,'for(i=0;i<params.length;i++){');
     UTL_FILE.PUT_LINE(fileHandler,'    if (value.toUpperCase() == params[i] ) {');
     UTL_FILE.PUT_LINE(fileHandler,'        field.style.backgroundColor = "white";');
     UTL_FILE.PUT_LINE(fileHandler,'        return true;');
     UTL_FILE.PUT_LINE(fileHandler,'    }');
     UTL_FILE.PUT_LINE(fileHandler,'  }');
     UTL_FILE.PUT_LINE(fileHandler,'  alert("Valid Values are: " +params);');
     UTL_FILE.PUT_LINE(fileHandler,'  field.focus();');
     UTL_FILE.PUT_LINE(fileHandler,'  field.style.backgroundColor = "cyan";');
     UTL_FILE.PUT_LINE(fileHandler,'  return false;');
     UTL_FILE.PUT_LINE(fileHandler,'}');

     UTL_FILE.PUT_LINE(fileHandler,'function validate_null(field, value) {');
     UTL_FILE.PUT_LINE(fileHandler,'if ( !value ) {');
     UTL_FILE.PUT_LINE(fileHandler,'alert("Enter field!");');
     UTL_FILE.PUT_LINE(fileHandler,'field.focus();');
     UTL_FILE.PUT_LINE(fileHandler,'field.style.backgroundColor = "cyan";');
     UTL_FILE.PUT_LINE(fileHandler,'return false;');
     UTL_FILE.PUT_LINE(fileHandler,'}');
     UTL_FILE.PUT_LINE(fileHandler,'else');
     UTL_FILE.PUT_LINE(fileHandler,'{');
     UTL_FILE.PUT_LINE(fileHandler,'field.style.backgroundColor = "white";');
     UTL_FILE.PUT_LINE(fileHandler,'return true;');
     UTL_FILE.PUT_LINE(fileHandler,'}');
     UTL_FILE.PUT_LINE(fileHandler,'}');
     UTL_FILE.PUT_LINE(fileHandler,'function validate_number(field, value) {');
     UTL_FILE.PUT_LINE(fileHandler,'if (parseFloat(value) != value )  {');
     UTL_FILE.PUT_LINE(fileHandler,'alert("Enter Valid Number!");');
     UTL_FILE.PUT_LINE(fileHandler,'field.focus();');
     UTL_FILE.PUT_LINE(fileHandler,'field.style.backgroundColor = "cyan";');
     UTL_FILE.PUT_LINE(fileHandler,'return false;');
     UTL_FILE.PUT_LINE(fileHandler,'}');
     UTL_FILE.PUT_LINE(fileHandler,'else');
     UTL_FILE.PUT_LINE(fileHandler,'{');
     UTL_FILE.PUT_LINE(fileHandler,'field.style.backgroundColor = "white";');
     UTL_FILE.PUT_LINE(fileHandler,'return true;');
     UTL_FILE.PUT_LINE(fileHandler,'}');
     UTL_FILE.PUT_LINE(fileHandler,'}');
     UTL_FILE.PUT_LINE(fileHandler,'function validate_date(field, value) {');
     UTL_FILE.PUT_LINE(fileHandler,'    try {');
     UTL_FILE.PUT_LINE(fileHandler,'        //Change the below values to determine which format of date you wish to check. It is set to mm/dd/yyyy by default.');
     UTL_FILE.PUT_LINE(fileHandler,'        var DayIndex = 1;');
     UTL_FILE.PUT_LINE(fileHandler,'        var MonthIndex = 0;');
     UTL_FILE.PUT_LINE(fileHandler,'        var YearIndex = 2;');

     UTL_FILE.PUT_LINE(fileHandler,'        value = value.replace(/-/g, "/").replace(/\./g, "/");');
     UTL_FILE.PUT_LINE(fileHandler,'        var SplitValue = value.split("/");');
     UTL_FILE.PUT_LINE(fileHandler,'        var OK = true;');
     UTL_FILE.PUT_LINE(fileHandler,'        if (!(SplitValue[DayIndex].length == 1 || SplitValue[DayIndex].length == 2)) {');
     UTL_FILE.PUT_LINE(fileHandler,'            OK = false;');
     UTL_FILE.PUT_LINE(fileHandler,'        }');
     UTL_FILE.PUT_LINE(fileHandler,'        if (OK && !(SplitValue[MonthIndex].length == 1 || SplitValue[MonthIndex].length == 2)) {');
     UTL_FILE.PUT_LINE(fileHandler,'            OK = false;');
     UTL_FILE.PUT_LINE(fileHandler,'        }');
     UTL_FILE.PUT_LINE(fileHandler,'        if (OK && SplitValue[YearIndex].length != 4) {');
     UTL_FILE.PUT_LINE(fileHandler,'            OK = false;');
     UTL_FILE.PUT_LINE(fileHandler,'        }');
     UTL_FILE.PUT_LINE(fileHandler,'        if (OK) {');
     UTL_FILE.PUT_LINE(fileHandler,'            var Day = parseInt(SplitValue[DayIndex], 10);');
     UTL_FILE.PUT_LINE(fileHandler,'            var Month = parseInt(SplitValue[MonthIndex], 10);');
     UTL_FILE.PUT_LINE(fileHandler,'            var Year = parseInt(SplitValue[YearIndex], 10);');

     UTL_FILE.PUT_LINE(fileHandler,'            if (OK = ((Year > 1900) && (Year <= new Date().getFullYear()))) {');
     UTL_FILE.PUT_LINE(fileHandler,'                if (OK = (Month <= 12 && Month > 0)) {');
     UTL_FILE.PUT_LINE(fileHandler,'                    var LeapYear = (((Year % 4) == 0) && ((Year % 100) != 0) || ((Year % 400) == 0));');

     UTL_FILE.PUT_LINE(fileHandler,'                    if (Month == 2) {');
     UTL_FILE.PUT_LINE(fileHandler,'                        OK = LeapYear ? Day<= 29 : Day <= 28;');
     UTL_FILE.PUT_LINE(fileHandler,'                    }');
     UTL_FILE.PUT_LINE(fileHandler,'                    else {');
     UTL_FILE.PUT_LINE(fileHandler,'                        if ((Month == 4) ||(Month == 6) || (Month == 9) || (Month == 11)) {');
     UTL_FILE.PUT_LINE(fileHandler,'                            OK = (Day > 0 && Day <= 30);');
     UTL_FILE.PUT_LINE(fileHandler,'                        }');
     UTL_FILE.PUT_LINE(fileHandler,'                        else {');
     UTL_FILE.PUT_LINE(fileHandler,'                            OK = (Day > 0 && Day <= 31);');
     UTL_FILE.PUT_LINE(fileHandler,'                        }');

     UTL_FILE.PUT_LINE(fileHandler,'                    }');
     UTL_FILE.PUT_LINE(fileHandler,'                }');
     UTL_FILE.PUT_LINE(fileHandler,'            }');
     UTL_FILE.PUT_LINE(fileHandler,'        }');
     UTL_FILE.PUT_LINE(fileHandler,'        if (  OK )');
     UTL_FILE.PUT_LINE(fileHandler,'        {');
     UTL_FILE.PUT_LINE(fileHandler,'            field.style.backgroundColor = "white";');
     UTL_FILE.PUT_LINE(fileHandler,'            return true;');
     UTL_FILE.PUT_LINE(fileHandler,'        }');
     UTL_FILE.PUT_LINE(fileHandler,'        else');
     UTL_FILE.PUT_LINE(fileHandler,'        {');
     UTL_FILE.PUT_LINE(fileHandler,'            alert("Invalid Date["+value+"].Enter Valid Date in format mm/dd/yyyy!");');
     UTL_FILE.PUT_LINE(fileHandler,'            field.focus();');
     UTL_FILE.PUT_LINE(fileHandler,'            field.style.backgroundColor = "cyan";');
     UTL_FILE.PUT_LINE(fileHandler,'            return false;');
     UTL_FILE.PUT_LINE(fileHandler,'        }');
     UTL_FILE.PUT_LINE(fileHandler,'    }');
     UTL_FILE.PUT_LINE(fileHandler,'    catch (e) {');
     UTL_FILE.PUT_LINE(fileHandler,'            alert("Invalid Date["+value+"].Enter Valid Date in format mm/dd/yyyy!");');
     UTL_FILE.PUT_LINE(fileHandler,'        field.focus();');
     UTL_FILE.PUT_LINE(fileHandler,'        field.style.backgroundColor = "cyan";');
     UTL_FILE.PUT_LINE(fileHandler,'        return false;');
     UTL_FILE.PUT_LINE(fileHandler,'    }');
     UTL_FILE.PUT_LINE(fileHandler,'}');

     UTL_FILE.PUT_LINE(fileHandler,'</script >');
     UTL_FILE.PUT_LINE(fileHandler,'</HEAD>');
     UTL_FILE.PUT_LINE(fileHandler,'        <%');
     UTL_FILE.PUT_LINE(fileHandler,'            sql = "SELECT  COUNT(1) CNT "');

     UTL_FILE.PUT_LINE(fileHandler,'            sql = sql + " FROM '||v_table_name||' A " ');
     UTL_FILE.PUT_LINE(fileHandler,'            Set objRs = oADOConnection.Execute(sql) ');
     UTL_FILE.PUT_LINE(fileHandler,'            allrecords = objRs.Fields("CNT").Value');
     FOR i IN 1..v_column_name_arr.LAST
     LOOP
         IF v_prev_column_name <> v_column_name_arr(i) THEN

            IF i = 1 THEN
               IF v_column_type_arr(i) = 'DATE' THEN
                  selectColStr := 'TO_CHAR(' || v_column_name_arr(i) ||',''MM/DD/YYYY'') ' || v_column_name_arr(i);
               ELSE
                  selectColStr := v_column_name_arr(i);
               END IF;
            ELSE
               IF v_column_type_arr(i) = 'DATE' THEN
                  selectColStr := selectColStr || ', TO_CHAR(' || v_column_name_arr(i) ||',''MM/DD/YYYY'') ' || v_column_name_arr(i) ;
               ELSE
                  selectColStr := selectColStr ||', ' || v_column_name_arr(i);
               END IF;

            END IF;
         END IF;
         v_prev_column_name := v_column_name_arr(i);

     END LOOP;
     UTL_FILE.PUT_LINE(fileHandler,'            sql = "SELECT  ' || selectColStr ||' FROM ' || v_table_name || ' A " ');
     UTL_FILE.PUT_LINE(fileHandler,'            Set oADORecordset = oADOConnection.Execute(sql)');

     UTL_FILE.PUT_LINE(fileHandler,'              for i=1 to offset ');
     UTL_FILE.PUT_LINE(fileHandler,'                 oADORecordset.movenext');
     UTL_FILE.PUT_LINE(fileHandler,'              Next');

     UTL_FILE.PUT_LINE(fileHandler,'        %>');

     UTL_FILE.PUT_LINE(fileHandler,'<BODY>');

     UTL_FILE.PUT_LINE(fileHandler,'<form name="myform" action="<%=ActionURL%>"method="post" onSubmit="">');

     UTL_FILE.PUT_LINE(fileHandler,'<TABLE BORDER="0" ALIGN="CENTER">');
     FOR i IN 1..v_column_name_arr.LAST
     LOOP
     -- Put Select Here
         IF v_column_type_arr(i) = 'DATE' THEN
            vDateFormatStr := ' (Date Format: mm/dd/yyyy)</B>';
         ELSE
            vDateFormatStr := '</B>';
         END IF;
         IF v_pk_constraint_type_arr(i) = 'P' THEN
            UTL_FILE.PUT_LINE(fileHandler,' <TR><TD><B>'||REPLACE(v_column_name_arr(i), '_', ' ')||vDateFormatStr||':</TD><TD><INPUT SIZE=50 NAME='||v_column_name_arr(i)||' READONLY TYPE=TEXT VALUE="<%=oADORecordset.Fields("'||v_column_name_arr(i)||'").Value%>" ></TD></TR>');
         ELSIF v_pk_constraint_type_arr(i) = 'R' THEN
            UTL_FILE.PUT_LINE(fileHandler,'        <%');
            UTL_FILE.PUT_LINE(fileHandler,'            sql = "SELECT DISTINCT  '||v_column_name_arr(i) || ' FROM ' || v_constraint_table_arr(i) ||' "');
            UTL_FILE.PUT_LINE(fileHandler,'            Set objRs'||v_column_name_arr(i)||' = oADOConnection.Execute(sql) ');
            UTL_FILE.PUT_LINE(fileHandler,'        %>');
            UTL_FILE.PUT_LINE(fileHandler,' <TR><TD><B>'||REPLACE(v_column_name_arr(i), '_', ' ')||vDateFormatStr||':</TD><TD><SELECT SIZE=5 NAME='||v_column_name_arr(i)||'  >');
            UTL_FILE.PUT_LINE(fileHandler,'        <%');
            UTL_FILE.PUT_LINE(fileHandler,'If objRs'||v_column_name_arr(i)||'.EOF Then' );
            UTL_FILE.PUT_LINE(fileHandler,'        %>');

            UTL_FILE.PUT_LINE(fileHandler,'    <OPTION VALUE="" ></OPTION>');

            UTL_FILE.PUT_LINE(fileHandler,'        <%');
            UTL_FILE.PUT_LINE(fileHandler,'Else');
            UTL_FILE.PUT_LINE(fileHandler,'  Do While Not objRs'||v_column_name_arr(i)||'.EOF');
            UTL_FILE.PUT_LINE(fileHandler,'        %>');
            UTL_FILE.PUT_LINE(fileHandler,'    <OPTION VALUE="<%=objRs'||v_column_name_arr(i)||'.Fields("'||v_column_name_arr(i) ||'").Value%>" <%if session("'|| REPLACE(INITCAP(v_column_name_arr(i)), '_') || 'Val") = objRs'||v_column_name_arr(i)||'.Fields("'||v_column_name_arr(i)||'").Value then  response.write "selected" end if%>> <%=objRs'||v_column_name_arr(i)||'.Fields("'||v_column_name_arr(i)||'").Value%></OPTION>');
            UTL_FILE.PUT_LINE(fileHandler,'        <%');
            UTL_FILE.PUT_LINE(fileHandler,'    objRs'||v_column_name_arr(i)||'.MoveNext()');
            UTL_FILE.PUT_LINE(fileHandler,'  Loop ');
            UTL_FILE.PUT_LINE(fileHandler,'End If');
            UTL_FILE.PUT_LINE(fileHandler,'objRs'||v_column_name_arr(i)||'.Close()');
            UTL_FILE.PUT_LINE(fileHandler,'        %>');
            UTL_FILE.PUT_LINE(fileHandler,' </SELECT></TD></TR>');
         ELSE -- If field is not Primary key or Foreign key
            validateFunctionCalls := NULL;
            IF v_pk_constraint_type_arr(i) = 'C' THEN -- If there is a constraint
               IF v_nullable_arr(i) = 'N' THEN -- If this is not null constraint
                  validateFunctionCalls := 'validate_null(this, this.value)';
               ELSE -- VALID VALUES CHECK Constraint
                  validateFunctionCalls := 'var list = ' || REPLACE (REPLACE(v_pk_search_condition_arr(i), v_column_name_arr(i) || ' IN (', '['), ')', ']') || '; validate_list(this, this.value, list) ';
               END IF;
            END IF;

            IF v_column_type_arr(i) = 'NUMBER' THEN
               IF validateFunctionCalls IS NULL THEN
                  validateFunctionCalls := 'validate_number(this, this.value)';
               ELSE
                  validateFunctionCalls := validateFunctionCalls || ';validate_number(this, this.value)';
               END IF;
            ELSIF v_column_type_arr(i) = 'DATE' THEN
               IF validateFunctionCalls IS NULL THEN
                  validateFunctionCalls := 'validate_date(this, this.value)';
               ELSE
                  validateFunctionCalls := validateFunctionCalls || ';validate_date(this, this.value)';
               END IF;
            END IF;

            IF validateFunctionCalls IS NULL THEN
               IF v_prev_column_name <> v_column_name_arr(i) THEN
                  UTL_FILE.PUT_LINE(fileHandler,' <TR><TD><B>'||REPLACE(v_column_name_arr(i), '_', ' ')||vDateFormatStr||':</TD><TD><INPUT SIZE=50 NAME='||v_column_name_arr(i)||' TYPE=TEXT VALUE="<%=oADORecordset.Fields("'||v_column_name_arr(i)||'").Value%>" ></TD></TR>');
               END IF;
            ELSE
               IF v_prev_column_name <> v_column_name_arr(i) THEN
                  UTL_FILE.PUT_LINE(fileHandler,' <TR><TD><B>'||REPLACE(v_column_name_arr(i), '_', ' ')||vDateFormatStr||':</TD><TD><INPUT SIZE=50 NAME='||v_column_name_arr(i)||' TYPE=TEXT onChange="'||validateFunctionCalls||'" VALUE="<%=oADORecordset.Fields("'||v_column_name_arr(i)||'").Value%>" ></TD></TR>');
               END IF;
            END IF;

         END IF;
         v_prev_column_name := v_column_name_arr(i);
     END LOOP;

     UTL_FILE.PUT_LINE(fileHandler,'<BR>');
     UTL_FILE.PUT_LINE(fileHandler,'<BR>');
     UTL_FILE.PUT_LINE(fileHandler,'<td><button align="CENTER"  type="button"><% if allrecords <> 1 then %>    <a href="'||v_tbl_pfx||'.asp?offset=<% = offset - 1%>">Prev Page</a><% else %>Prev Page<% end if %></button></td>');
     UTL_FILE.PUT_LINE(fileHandler,'<td><input type="submit" name="Insert" value="Insert" onclick="javascript:if(!confirm(''This action will Insert a new record. Are you sure?'')){return false;}else{form.action='''||v_tbl_pfx||'.asp?Insert=TRUE'';}"></Input></td>');
     UTL_FILE.PUT_LINE(fileHandler,'<td><input type="submit" name="Update" value="Update" onclick="javascript:if(!confirm(''This action will Update the currentrecord. Are you sure?'')){return false;}else{form.action='''||v_tbl_pfx||'.asp?Update=TRUE'';}"></Input></td>');
     UTL_FILE.PUT_LINE(fileHandler,'<td><input type="submit" name="Delete" value="Delete" onclick="javascript:if(!confirm(''This action will Delete the currentrecord. Are you sure?'')){return false;}else{form.action='''||v_tbl_pfx||'.asp?Delete=TRUE'';}"></Input></td>');
     UTL_FILE.PUT_LINE(fileHandler,'<td><button align="CENTER"  type="button"><% if allrecords <> 1 and allrecords > offset+1 then %>    <a href="'||v_tbl_pfx||'.asp?offset=<% = offset + 1%>">Next Page</a><% else %>Last Record<% end if %></button></td>');

     UTL_FILE.PUT_LINE(fileHandler,'</table>');
     UTL_FILE.PUT_LINE(fileHandler,'</form>');

     UTL_FILE.PUT_LINE(fileHandler,'<tr>');

     UTL_FILE.PUT_LINE(fileHandler,'</tr>');
     UTL_FILE.PUT_LINE(fileHandler,'</BODY>');
     UTL_FILE.PUT_LINE(fileHandler,'</HTML>');
     UTL_FILE.PUT_LINE(fileHandler,'<%');
     UTL_FILE.PUT_LINE(fileHandler,'oADORecordset.Close');
     UTL_FILE.PUT_LINE(fileHandler,'objRs.Close');
     UTL_FILE.PUT_LINE(fileHandler,'oADOConnection.Close');
     UTL_FILE.PUT_LINE(fileHandler,'%>');
     UTL_FILE.FCLOSE(fileHandler);
END;
--#################
PROCEDURE create_ASP_list_page_prc(v_table_name IN VARCHAR2)
IS
v_tbl_pfx VARCHAR2(30);
v_pkg_name VARCHAR2(30);
c cur_typ;
sqlStr VARCHAR2(1000);
validateFunctionCalls VARCHAR2(1000);
TYPE column_name_arr IS VARRAY(200) OF all_tab_columns.column_name%TYPE;
TYPE pk_constraint_type_arr IS VARRAY(200) OF all_constraints.constraint_type%TYPE;
TYPE column_type_arr IS VARRAY(200) OF all_tab_columns.data_type%TYPE;
TYPE constraint_table_arr IS VARRAY(200) OF all_tab_columns.table_name%TYPE;
TYPE nullable_arr IS VARRAY(200) OF all_tab_columns.nullable%TYPE;
TYPE pk_search_condition_arr IS VARRAY(200) OF all_constraints.search_condition%TYPE;
v_column_name_arr column_name_arr := column_name_arr();
v_column_type_arr column_type_arr := column_type_arr();
v_pk_constraint_type_arr pk_constraint_type_arr := pk_constraint_type_arr();
v_constraint_table_arr constraint_table_arr := constraint_table_arr();
v_nullable_arr nullable_arr := nullable_arr();
v_pk_search_condition_arr pk_search_condition_arr := pk_search_condition_arr();
fileHandler UTL_FILE.FILE_TYPE;
vDateFormatStr VARCHAR2(100):= '</B>';
j NUMBER :=0;
whereClause VARCHAR2(1000);
logical_delete_found NUMBER:= 0;
CURSOR all_fields_cur(v_tbl_name IN VARCHAR2)
IS
SELECT c.column_name, c.data_type,pk.constraint_type, DECODE(pk.constraint_type, 'R', (SELECT DISTINCT table_name
       FROM all_constraints
       WHERE constraint_name = pk.r_constraint_name),
       NULL) constraint_table, nullable, pk.search_condition
FROM   all_tab_columns c,
(     SELECT a.table_name, b.column_name,a.constraint_type, a.r_constraint_name, a.search_condition
     FROM   all_constraints a,
            all_cons_columns  b
     WHERE a.constraint_name = b.constraint_name )pk
WHERE  c.table_name = pk.table_name(+)
AND    c.column_name = pk.column_name(+)
AND    c.table_name = v_tbl_name
ORDER  BY c.column_id ASC;

BEGIN
     v_column_name_arr.EXTEND;
     v_column_type_arr.EXTEND;
     v_pk_constraint_type_arr.EXTEND;
     v_constraint_table_arr.EXTEND;
     v_nullable_arr.EXTEND;
     v_pk_search_condition_arr.EXTEND;

     SELECT REPLACE(UPPER(v_table_name), '_TBL')||'_PKG',REPLACE(lower(v_table_name), '_tbl')
     INTO v_pkg_name, v_tbl_pfx
     FROM DUAL;
     fileHandler := UTL_FILE.FOPEN('ASP_DIR', UPPER(v_tbl_pfx) || '.asp','W');
     -- GET All columns
     OPEN all_fields_cur(v_table_name);
     FETCH all_fields_cur BULK COLLECT INTO v_column_name_arr ,
                                            v_column_type_arr ,
                                            v_pk_constraint_type_arr,
                                            v_constraint_table_arr,
                                            v_nullable_arr,
                                            v_pk_search_condition_arr;
     UTL_FILE.PUT_LINE(fileHandler,'<%@ Language=VBScript %>');
     UTL_FILE.PUT_LINE(fileHandler,'<%');

     UTL_FILE.PUT_LINE(fileHandler,'connstr = "DSN=ODBCTESTDB;UID=SCOTT;PWD=TIGER"');
     UTL_FILE.PUT_LINE(fileHandler,'Set oADOConnection = Server.CreateObject("ADODB.Connection")');
     UTL_FILE.PUT_LINE(fileHandler,'oADOConnection.Provider = "MSDASQL"');
     UTL_FILE.PUT_LINE(fileHandler,'oADOConnection.ConnectionString = connstr');

     UTL_FILE.PUT_LINE(fileHandler,'oADOConnection.CommandTimeout = 0');
     UTL_FILE.PUT_LINE(fileHandler,'oADOConnection.Open');
     UTL_FILE.PUT_LINE(fileHandler,'%>');
     UTL_FILE.PUT_LINE(fileHandler,'<html>');
     UTL_FILE.PUT_LINE(fileHandler,'<head>');
     UTL_FILE.PUT_LINE(fileHandler,'<H1 ALIGN="CENTER"><B>'||REPLACE(v_table_name,'_', ' ') ||'</B></H1>');
     UTL_FILE.PUT_LINE(fileHandler,'<meta http-equiv="Content-Type"');
     UTL_FILE.PUT_LINE(fileHandler,'content="text/html; charset=iso-8859-1">');
     UTL_FILE.PUT_LINE(fileHandler,'<meta name="GENERATOR" content="Microsoft FrontPage Express 2.0">');
     UTL_FILE.PUT_LINE(fileHandler,'<title>'||REPLACE(v_table_name,'_', ' ') ||'</title>');
     UTL_FILE.PUT_LINE(fileHandler,'<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">');
     UTL_FILE.PUT_LINE(fileHandler,'<link href="style.css" rel="stylesheet" type="text/css">');
     UTL_FILE.PUT_LINE(fileHandler,'<style type="text/css">');
     UTL_FILE.PUT_LINE(fileHandler,'<!--');
     UTL_FILE.PUT_LINE(fileHandler,'body {');
     UTL_FILE.PUT_LINE(fileHandler,'    background-color: #FFFFFF;');
     UTL_FILE.PUT_LINE(fileHandler,'}');
     UTL_FILE.PUT_LINE(fileHandler,'-->');
     UTL_FILE.PUT_LINE(fileHandler,'</style>');
     UTL_FILE.PUT_LINE(fileHandler,'</head>');
     UTL_FILE.PUT_LINE(fileHandler,'<body class="mainbody">');
     UTL_FILE.PUT_LINE(fileHandler,'<form name="myform" action="<%=ActionURL%>"method="post" onSubmit="">');
     UTL_FILE.PUT_LINE(fileHandler,'<table border=1 cellspacing=1 cellpadding=1width="100%">');
     UTL_FILE.PUT_LINE(fileHandler,'  <tr>');
     UTL_FILE.PUT_LINE(fileHandler,'  <th align="left" bgcolor="#C0C0C0" bordercolor="#FFFFFF">SERIAL#</th>');
     FOR i IN 1..v_column_name_arr.LAST
     LOOP
         UTL_FILE.PUT_LINE(fileHandler,'  <th align="left" bgcolor="#C0C0C0" bordercolor="#FFFFFF">'||REPLACE(v_column_name_arr(i), '_', ' ') ||'</th>');
     END LOOP;
     UTL_FILE.PUT_LINE(fileHandler,'  </tr>');
     UTL_FILE.PUT_LINE(fileHandler,'<%');
     UTL_FILE.PUT_LINE(fileHandler,'    sql = "SELECT ROWNUM serial# "');
     FOR i IN 1..v_column_name_arr.LAST
     LOOP
         UTL_FILE.PUT_LINE(fileHandler,'    sql = sql + ", '||v_column_name_arr(i)||'"');
     END LOOP;
     UTL_FILE.PUT_LINE(fileHandler,'    sql = sql + " FROM   ' ||v_table_name ||' " ');
     UTL_FILE.PUT_LINE(fileHandler,'    sql = sql + " ORDER BY 1 "');
     UTL_FILE.PUT_LINE(fileHandler,'    Set oADORecordset = oADOConnection.Execute(sql)');
     UTL_FILE.PUT_LINE(fileHandler,'    Do While Not oADORecordset.EOF');
     UTL_FILE.PUT_LINE(fileHandler,'%>');

     UTL_FILE.PUT_LINE(fileHandler,'  <tr>');
     UTL_FILE.PUT_LINE(fileHandler,'    <td nowrap><div align=right><%=oADORecordset.Fields("SERIAL#").Value%></div></td>');
     FOR i IN 1..v_column_name_arr.LAST
     LOOP
         UTL_FILE.PUT_LINE(fileHandler,'    <td nowrap><div align=right><%=oADORecordset.Fields("'||v_column_name_arr(i) || '").Value%></div></td>');
     END LOOP;
     UTL_FILE.PUT_LINE(fileHandler,'  </tr>');
     UTL_FILE.PUT_LINE(fileHandler,'<%');
     UTL_FILE.PUT_LINE(fileHandler,'   oADORecordset.MoveNext');
     UTL_FILE.PUT_LINE(fileHandler,'   Loop');
     UTL_FILE.PUT_LINE(fileHandler,'   oADORecordset.Close');
     UTL_FILE.PUT_LINE(fileHandler,'%>');
     UTL_FILE.PUT_LINE(fileHandler,'</table>');
     UTL_FILE.PUT_LINE(fileHandler,'</form>');
     UTL_FILE.PUT_LINE(fileHandler,'</body>');
     UTL_FILE.PUT_LINE(fileHandler,'</html>');

     UTL_FILE.FCLOSE(fileHandler);
END;
