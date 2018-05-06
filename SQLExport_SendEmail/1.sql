--Date : 11-Sept-2015
--Version 1.0
--Developed by Nishad Bhandarkar

column tm new_value file_time noprint
select to_char(CURRENT_DATE, 'DDMMYYYY') tm from dual;

set linesize 32767  LONG 500 LONGCHUNKSIZE 500
SET COLSEP "	"
SET FEEDBACK OFF
SET NEWPAGE NONE
SET VERIFY OFF
SET HEADING ON
SET WRAP OFF
TTITLE OFF
BTITLE OFF
SET TERMOUT OFF
SET ECHO OFF
SET TRIMSPOOL ON
SET newpage none 
SET trims ON 
set underline off
set pagesize 50000
buffer=10000000

set markup html on spool on
SPOOL <FilenameGoesHere>_&file_time..xls;

<<Your SQL query goes here>>
 
set markup html off spool off
SPOOL OFF

SET ECHO ON
SET LINESIZE 80
SET COLSEP " "
SET FEEDBACK ON
SET NEWPAGE 12
SET TERMOUT ON
SET VERIFY ON
SET WRAP ON
SET HEADING ON
TTITLE ON
BTITLE ON
EXIT
/ 
 