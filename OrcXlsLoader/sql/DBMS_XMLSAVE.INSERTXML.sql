declare
    targetTable varchar2(30):=?;
    insctx		dbms_xmlsave.ctxtype;
	n_rows		number;
	s_xml		clob;
begin
    
    insctx := dbms_xmlsave.newContext(targetTable);

    dbms_xmlsave.setdateformat(insctx, 'dd/MM/yyyy HH:mm:ss');
    execute immediate 'delete '||targetTable;    
    s_xml := ?;
    n_rows := dbms_xmlsave.insertXML(insctx, s_xml);
    
exception
    when others then
        dbms_xmlsave.closeContext(insctx);
        raise;
end;
