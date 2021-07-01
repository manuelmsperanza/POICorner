select ui.table_name, ui.index_name, uic.column_name, uic.DESCEND, uic.COLUMN_POSITION
from user_indexes ui, user_ind_columns uic 
where ui.index_type <> 'LOB'
and ui.index_name = uic.INDEX_NAME (+)
and not exists (select 1 from user_constraints uc where uc.index_name = ui.index_name)
order by table_name, index_name, COLUMN_POSITION