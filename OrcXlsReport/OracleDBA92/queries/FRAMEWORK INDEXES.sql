select ui.table_name, ui.index_name, list_cols.column_list
from user_indexes ui, (select TABLE_NAME, INDEX_NAME,  LISTAGG(column_name||' '||DESCEND,', ') WITHIN GROUP (ORDER BY COLUMN_POSITION) column_list
from user_ind_columns
group by table_name, index_name) list_cols
where ui.index_type <> 'LOB'
and ui.index_name = list_cols.INDEX_NAME (+)
and not exists (select 1 from user_constraints uc where uc.index_name = ui.index_name)