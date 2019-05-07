#Create a new project
mvn archetype:generate -DarchetypeCatalog=org.apache.maven.archetypes -Dfilter=maven-archetype-quickstart -DgroupId=me.hoffnungland -DartifactId=OrcXlsLoader -Dpackage=me.hoffnungland.poi.corner.orcxlsloader -Dversion=0.0.1-SNAPSHOT
#Build settings
##Remove junit:junit:3.8.1

#how fill properties files
	select COLUMN_NAME||'='||COLUMN_NAME
	from cols where table_name = <<table_name>> order by column_id
	/
	select COLUMN_NAME||'.type='||DATA_TYPE
	from cols where table_name = <<table_name>> order by column_id
	/


#add .gitignore to mandatory empty directory
	# Ignore everything in this directory
	*
	# Except this file
	!.gitignore
