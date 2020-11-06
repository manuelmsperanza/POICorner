declare 
	v_procedure_name varchar2(30) := ?;
begin

	dbms_scheduler.create_job(
		job_name	=>	'One_Time_Job_'||v_procedure_name,  
		job_type	=>	'STORED_PROCEDURE',  
		job_action	=>	v_procedure_name,  
		start_date	=>	sysdate,  
		enabled		=>	TRUE,  
		auto_drop	=>	TRUE,  
		comments	=>	'one-time job');

end;