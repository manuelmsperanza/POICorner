#Create a new project
	mvn archetype:generate -Dfilter="org.apache.maven.archetypes:maven-archetype-quickstart" -DgroupId="com.hoffnungland" -DartifactId=JSONXlsLoader -Dpackage="com.hoffnungland.poi.corner.jsonxlsloader" -Dversion="0.0.1-SNAPSHOT"
#Build settings

#add .gitignore to mandatory empty directory
	# Ignore everything in this directory
	*
	# Except this file
	!.gitignore
