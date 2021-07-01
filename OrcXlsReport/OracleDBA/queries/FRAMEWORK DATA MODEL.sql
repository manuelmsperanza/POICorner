select a.TABLE_NAME, COLUMN_NAME, decode(data_type, 'VARCHAR2', data_type||'('||DATA_LENGTH||')', data_type) DATA_TYPE, 
null "Sequence", null "Object_Type_ID", null "Functional_Field",
DECODE(data_type, 'BLOB', null, 'DATE', 'java.util.Date', 'NUMBER', 'java.lang.Long', 'TIMESTAMP', 'java.util.Date', 'java.lang.String') "Hibernate data type mapping",
DECODE(data_type, 'CLOB', 'Y') "XML Valid"
from user_tab_columns a, user_tables b
where a.table_name = b.table_name
order by table_name, column_id