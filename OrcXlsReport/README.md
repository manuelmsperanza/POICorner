# OrcXlsReport

Oracle Excel report executes queries into the database and printout data into the worksheets.
Run the main method in com.hoffnungland.poi.corner.orcxlsreport.App with the following parameters:

* ConnectionName
* ProjectName
* ExcelName
* TargetPath
* id
* name

ConnectionName and ProjectName configurations are detailed below.

## Property files
Connection file name must be in etc/connections directory and has the same rules listed in OracleConn.

## Projects
There are different type of queries you can apply. Each group of queries is stored into a directory.
The name of the directory corresponds to the input parameter

* metadata
* queries
* queriesById
* queriesByName
* queriesJnt
* queriesJntCached
* snapshot
* tables

## Create a new project
	mvn archetype:generate -Dfilter="org.apache.maven.archetypes:maven-archetype-quickstart" -DgroupId="com.hoffnungland" -DartifactId=OrcXlsReport -Dpackage="com.hoffnungland.poi.corner.orcxlsreport" -Dversion="0.0.1-SNAPSHOT"
#Build settings
##Remove junit:junit:3.8.1

#add .gitignore to mandatory empty directory
	# Ignore everything in this directory
	*
	# Except this file
	!.gitignore
