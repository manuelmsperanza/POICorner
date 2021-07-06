# ExtractShapesVsdx

Read the *.vsdx files and printout the text of each shapes with the name of the sheet it belongs.

## Create a new project
	mvn archetype:generate -Dfilter="org.apache.maven.archetypes:maven-archetype-quickstart" -DgroupId="com.hoffnungland" -DartifactId=ExtractShapesVsdx -Dpackage="com.hoffnungland.poi.corner.extractshapesvsdx" -Dversion="0.0.1-SNAPSHOT"
## Build settings
### Remove junit:junit:3.8.1


#add .gitignore to mandatory empty directory
	# Ignore everything in this directory
	*
	# Except this file
	!.gitignore
