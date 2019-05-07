#Create a new project
mvn archetype:generate -DarchetypeCatalog=org.apache.maven.archetypes -Dfilter=maven-archetype-quickstart -DgroupId=me.hoffnungland -DartifactId=OrcXlsReport -Dpackage=me.hoffnungland.poi.corner.orcxlsreport -Dversion=0.0.1-SNAPSHOT
#Build settings
##Remove junit:junit:3.8.1

#add .gitignore to mandatory empty directory
	# Ignore everything in this directory
	*
	# Except this file
	!.gitignore
