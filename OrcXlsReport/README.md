# OrcXlsReport

Oracle Excel report executes queries into the database and printout data into the worksheets.
Run the main method in com.hoffnungland.poi.corner.orcxlsreport.App with the following parameters:

* ConnectionName: the target environment
* ProjectName: the directory containing the queries.
* ExcelName: the output excel name (except for queries in metadata, snapshot and tables).
* TargetPath: the target directory will contain the workbooks
* id: optional parameter used by queriesById (hint: use 0 or a negative number to skip these queries)
* name: optional parameter used by queriesByName

ConnectionName and ProjectName configurations are detailed below.

## Property files
Connection file name must be in etc/connections directory and has the same rules listed in OracleConn.

## Projects
Project directories must in etc directory. The ProjectName input parameter is the name of a directory you intend use.
Each project will contains the following sub-directories:

* metadata: contains a list of \*.txt files (one per output workbook - name will be the same of the \*.txt file). Into each file you must list the tables you intend query (a worksheet per line)
* queries: contains a list \*.sql files (one per worksheet in ExcelName - name will be the same of the \*.sql). Each file contains the query to run (without end of statement)
* queriesById: contains a list \*.sql files (one per worksheet in ExcelName - name will be the same of the \*.sql). Each file contains the query to run (without end of statement) with a variable bound to id input parameter.
* queriesByName: contains a list \*.sql files (one per worksheet in ExcelName - name will be the same of the \*.sql). Each file contains the query to run (without end of statement) with a variable bound to name input parameter.
* queriesJnt: contains a list \*.sql files (one per worksheet in ExcelName - name will be the same of the \*.sql). Each file contains the query to run (without end of statement) having two columns STM and JUNCTION.
The output of this query will be used to generate another query (joining STM and JUNCTION of each row) with the actual output. JUNCTION is a set operator (union, union all...).
* queriesJntCached: the same of queriesJnt but cached
* snapshot: the same of metadata but not cached.
* tables: the same of metadata, but the header is simpler (has only one row containing only the column list, without the first line containing the name of the table).


## Create a new project
	mvn archetype:generate -Dfilter="org.apache.maven.archetypes:maven-archetype-quickstart" -DgroupId="com.hoffnungland" -DartifactId=OrcXlsReport -Dpackage="com.hoffnungland.poi.corner.orcxlsreport" -Dversion="0.0.1-SNAPSHOT"
#Build settings
##Remove junit:junit:3.8.1

#add .gitignore to mandatory empty directory
	# Ignore everything in this directory
	*
	# Except this file
	!.gitignore
