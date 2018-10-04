#Create a new project
mvn archetype:generate -DarchetypeCatalog=http://repo.maven.apache.org/maven2/archetype-catalog.xml -Dfilter=maven-archetype-quickstart -DgroupId=me.hoffnungland -DartifactId=XmlXlsReport -Dpackage=me.hoffnungland.poi.corner.xmlxlsreport -Dversion=0.0.1-SNAPSHOT
#Build settings
##Remove junit:junit:3.8.1


#add .gitignore to mandatory empty directory
	# Ignore everything in this directory
	*
	# Except this file
	!.gitignore
