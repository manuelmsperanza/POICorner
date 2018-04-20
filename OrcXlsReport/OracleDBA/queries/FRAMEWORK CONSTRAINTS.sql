with list_cols as (select ucc.table_name, ucc.constraint_name, LISTAGG(ucc.column_name,', ') WITHIN GROUP (ORDER BY ucc.position) column_list
from user_cons_columns ucc
group by ucc.table_name, ucc.constraint_name)
select list_cols.table_name, DECODE(uc.constraint_type, 'P', 'PK', 'U', 'UNIQUE', '?')  CONSTRAINT_TYPE, list_cols.constraint_name, list_cols.column_list, uc.R_CONSTRAINT_NAME column_list_ref, uc.SEARCH_CONDITION, uc.DEFERRABLE, uc.DEFERRED, uc.VALIDATED
from user_constraints uc, list_cols
where uc.constraint_type in('P', 'U')
and uc.constraint_name = list_cols.constraint_name
UNION ALL
select list_cols.table_name, 'CHECK' CONSTRAINT_TYPE, list_cols.constraint_name, list_cols.column_list, uc.R_CONSTRAINT_NAME column_list_ref, uc.SEARCH_CONDITION, uc.DEFERRABLE, uc.DEFERRED, uc.VALIDATED
from user_constraints uc, list_cols
where uc.constraint_type = 'C'
and uc.constraint_name = list_cols.constraint_name
UNION ALL
select a.table_name, 'FK' CONSTRAINT_TYPE, a.constraint_name, a.column_list, b.table_name||' ('||b.column_list||')' column_list_ref, uc.SEARCH_CONDITION, uc.DEFERRABLE, uc.DEFERRED, uc.VALIDATED
from user_constraints uc, list_cols a, list_cols b
where uc.constraint_type = 'R'
and uc.constraint_name = a.constraint_name
and uc.R_CONSTRAINT_NAME = b.constraint_name
order by table_name, constraint_name