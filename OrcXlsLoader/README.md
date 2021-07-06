# OrcXlsLoader

Oracle excel loader can help you to bulk load the content of an excel file straight into a database.
Each sheet can represent a database table.
To make it work, you need to create a property file for every sheet and every environment you intend deploy your data.

## Property files

### Connection property file
Connection file name must be in etc/connections directory and has the same rules listed in OracleConn.

### Data mapping property file
Data mapping file must be in etc/worksheet_name/ directory and has the following naming convention: datamapping.connection_name.properties.

The file contains these properties

* TABLE_NAME:
* TABLE.backup:
* TABLE.clean:
* EXEC_POST_LOAD:
* SCHEDULE_PROCEDURE_POST_LOAD:


#### How fill data mapping properties files

You can run the following queries, copy the result and paste in the target file.

	select COLUMN_NAME||'='||COLUMN_NAME
	from cols where table_name = <<table_name>> order by column_id
	/
	select COLUMN_NAME||'.type='||DATA_TYPE
	from cols where table_name = <<table_name>> order by column_id
	/

## Create a new project
	mvn archetype:generate -Dfilter="org.apache.maven.archetypes:maven-archetype-quickstart" -DgroupId="com.hoffnungland" -DartifactId=OrcXlsLoader -Dpackage="com.hoffnungland.poi.corner.orcxlsloader" -Dversion="0.0.1-SNAPSHOT"
	
## Build settings
### Remove junit:junit:3.8.1




#add .gitignore to mandatory empty directory
	# Ignore everything in this directory
	*
	# Except this file
	!.gitignore
