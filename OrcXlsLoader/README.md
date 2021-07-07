# OrcXlsLoader

Oracle Excel loader can help you to bulk load the content of an excel file straight into a database.
Each sheet can represent a database table.
To make it work, you need to create a property file for every sheet and every environment you intend deploy your data.

## Property files

### Connection property file
Connection file name must be in etc/connections directory and has the same rules listed in OracleConn.

### Data mapping property file
Data mapping file must be in etc/worksheet_name/ directory and has the following naming convention: datamapping.connection_name.properties.

The file contains these properties

* TABLE_NAME: the target table name
* TABLE.backup: create an excel file containing the data before loading. 
* TABLE.clean: if true, delete all rows before doing the insert
* EXEC_POST_LOAD: procedure or function to invoke suddenly after the loading
* SCHEDULE_PROCEDURE_POST_LOAD: procedure or function to scheduler after the loading
* list of header_column_name mapped with the related table_column_name (e.g. Header #1=COLUMN_A)
* list of table_column_name.type mapped with the related data type (e.g. COLUMN_A.type=VARCHAR2)

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


# Run with Maven
	
	start mvn exec:java -Dexec.mainClass="com.hoffnungland.poi.corner.orcxlsloader.App" -Dlog4j.configurationFile=src/main/resources/log4j2.xml

# Create Jar with dependencies

## Configure the pom.xml

	<plugin>
		<artifactId>maven-assembly-plugin</artifactId>
		<configuration>
			<descriptorRefs>
				<descriptorRef>jar-with-dependencies</descriptorRef>
			</descriptorRefs>
			<appendAssemblyId>false</appendAssemblyId>
			<finalName>${project.artifactId}</finalName>
			<archive>
				<manifest>
					<mainClass>com.hoffnungland.poi.corner.orcxlsloader.App</mainClass>
				</manifest>
			</archive>
		</configuration>
	</plugin>

## Execute the maven assembly single

	mvn assembly:single

#add .gitignore to mandatory empty directory
	# Ignore everything in this directory
	*
	# Except this file
	!.gitignore
