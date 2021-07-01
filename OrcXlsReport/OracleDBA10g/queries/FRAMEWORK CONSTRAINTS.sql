select ucc.table_name, DECODE(uc.constraint_type, 'P', 'PK', 'U', 'UNIQUE', '?')  CONSTRAINT_TYPE, ucc.constraint_name, ucc.column_name, uc.R_CONSTRAINT_NAME column_list_ref, uc.SEARCH_CONDITION, uc.DEFERRABLE, uc.DEFERRED, uc.VALIDATED
from user_constraints uc, user_cons_columns ucc
where uc.constraint_type in('P', 'U')
and uc.constraint_name = ucc.constraint_name
UNION ALL
select ucc.table_name, 'CHECK' CONSTRAINT_TYPE, ucc.constraint_name, ucc.column_name, uc.R_CONSTRAINT_NAME column_list_ref, uc.SEARCH_CONDITION, uc.DEFERRABLE, uc.DEFERRED, uc.VALIDATED
from user_constraints uc, user_cons_columns ucc
where uc.constraint_type = 'C'
and uc.constraint_name = ucc.constraint_name
UNION ALL
select ucca.table_name, 'FK' CONSTRAINT_TYPE, ucca.constraint_name, ucca.column_name, uccb.table_name||' ('||uccb.column_name||')' column_list_ref, uc.SEARCH_CONDITION, uc.DEFERRABLE, uc.DEFERRED, uc.VALIDATED
from user_constraints uc, user_cons_columns ucca, user_cons_columns uccb 
where uc.constraint_type = 'R'
and uc.constraint_name = ucca.constraint_name
and uc.R_CONSTRAINT_NAME = uccb.constraint_name
and ucca.position = uccb.position
order by table_name, constraint_name