#Create a new project
mvn archetype:generate -Dfilter="org.apache.maven.archetypes:maven-archetype-quickstart" -DgroupId="com.hoffnungland" -DartifactId=XmlXlsReport -Dpackage="com.hoffnungland.poi.corner.xmlxlsreport" -Dversion="0.0.1-SNAPSHOT"
#Build settings
##Remove junit:junit:3.8.1


#add .gitignore to mandatory empty directory
	# Ignore everything in this directory
	*
	# Except this file
	!.gitignore
