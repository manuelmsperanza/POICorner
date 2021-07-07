# ExtractShapesVsdx

Read the *.vsdx files and printout the text of each shapes with the name of the sheet it belongs.

## Create a new project
	mvn archetype:generate -Dfilter="org.apache.maven.archetypes:maven-archetype-quickstart" -DgroupId="com.hoffnungland" -DartifactId=ExtractShapesVsdx -Dpackage="com.hoffnungland.poi.corner.extractshapesvsdx" -Dversion="0.0.1-SNAPSHOT"
## Build settings
### Remove junit:junit:3.8.1


# Run with Maven
	
	start mvn exec:java -Dexec.mainClass="com.hoffnungland.poi.corner.extractshapesvsdx.App" -Dlog4j.configurationFile=src/main/resources/log4j2.xml

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
					<mainClass>com.hoffnungland.poi.corner.extractshapesvsdx.App</mainClass>
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
