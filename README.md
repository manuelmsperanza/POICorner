#Create a new project
mvn archetype:generate -DarchetypeCatalog=http://repo.maven.apache.org/maven2/archetype-catalog.xml -Dfilter=maven-archetype-quickstart -DgroupId=***REMOVED*** -DartifactId=POICorner -Dpackage=***REMOVED***.poi.corner -Dversion=0.0.1-SNAPSHOT
#Build settings
## Delete the src directory
## Change the package type

	<packaging>pom</packaging>

##Add prerequisites

	<prerequisites>
		<maven>3.0</maven>
	</prerequisites>
##Configure the compiler
Update to java 1.7<br>

	<properties>
		<project.build.sourceEncoding>UTF-8</project.build.sourceEncoding>
		<java.source.version>1.7</java.source.version>
		<java.target.version>1.7</java.target.version>
	</properties>
	<build>
		<pluginManagement>
			<plugins>
				<plugin>
					<groupId>org.apache.maven.plugins</groupId>
					<artifactId>maven-compiler-plugin</artifactId>
					<!-- version>3.5.1</version -->
					<configuration>
						<encoding>UTF-8</encoding>
						<source>${java.source.version}</source>
						<target>${java.target.version}</target>
					</configuration>
				</plugin>
			</plugins>
		</pluginManagement>
	</build>

#Relationship
##Add the dependencies
###Oracle jdbc dependencies
[Add the Oracle Maven Repository](http://docs.oracle.com/middleware/1213/core/MAVEN/config_maven_repo.htm#MAVEN9010)
###Instruction to encrypt the password on maven settings.xml
[Encryption guide](http://maven.apache.org/guides/mini/guide-encryption.html)<br>
Add log4j, jdbc e POI update jUnit<br>


	<dependencyManagement>
		<dependencies>
			<dependency>
				<groupId>org.apache.logging.log4j</groupId>
				<artifactId>log4j-bom</artifactId>
				<version>2.6.2</version>
				<scope>import</scope>
				<type>pom</type>
			</dependency>
		</dependencies>
	</dependencyManagement>
	<dependencies>
		<!-- https://mvnrepository.com/artifact/junit/junit -->
		<dependency>
			<groupId>junit</groupId>
			<artifactId>junit</artifactId>
			<version>4.12</version>
			<scope>test</scope>
		</dependency>
		<!-- https://maven.oracle.com -->
		<dependency>
			<groupId>com.oracle.jdbc</groupId>
			<artifactId>ojdbc7</artifactId>
			<version>12.1.0.2</version>
		</dependency>
		<!-- https://mvnrepository.com/artifact/org.apache.poi/poi -->
		<dependency>
			<groupId>org.apache.poi</groupId>
			<artifactId>poi</artifactId>
			<version>${poi.target.version}</version>
		</dependency>
		<!-- https://mvnrepository.com/artifact/org.apache.poi/poi-ooxml -->
		<dependency>
			<groupId>org.apache.poi</groupId>
			<artifactId>poi-ooxml</artifactId>
			<version>${poi.target.version}</version>
		</dependency>
		<!-- https://mvnrepository.com/artifact/org.apache.poi/poi-ooxml-schemas -->
		<dependency>
			<groupId>org.apache.poi</groupId>
			<artifactId>poi-ooxml-schemas</artifactId>
			<version>${poi.target.version}</version>
		</dependency>
		<!-- https://mvnrepository.com/artifact/org.apache.poi/poi-scratchpad -->
		<dependency>
			<groupId>org.apache.poi</groupId>
			<artifactId>poi-scratchpad</artifactId>
			<version>${poi.target.version}</version>
		</dependency>
		<!-- https://mvnrepository.com/artifact/org.apache.poi/ooxml-schemas -->
		<dependency>
			<groupId>org.apache.poi</groupId>
			<artifactId>ooxml-schemas</artifactId>
			<version>1.3</version>
		</dependency>
		<!-- https://mvnrepository.com/artifact/org.apache.poi/poi-contrib -->
		<dependency>
			<groupId>org.apache.poi</groupId>
			<artifactId>poi-contrib</artifactId>
			<version>3.6</version>
		</dependency>
		<!-- https://mvnrepository.com/artifact/org.apache.poi/poi-excelant -->
		<dependency>
			<groupId>org.apache.poi</groupId>
			<artifactId>poi-excelant</artifactId>
			<version>${poi.target.version}</version>
		</dependency>
		<!-- https://mvnrepository.com/artifact/org.apache.poi/ooxml-security -->
		<dependency>
			<groupId>org.apache.poi</groupId>
			<artifactId>ooxml-security</artifactId>
			<version>1.1</version>
		</dependency>
		<!-- https://mvnrepository.com/artifact/org.apache.logging.log4j/log4j-api -->
		<dependency>
			<groupId>org.apache.logging.log4j</groupId>
			<artifactId>log4j-api</artifactId>
			<!--version>2.6.1</version -->
		</dependency>
		<!-- https://mvnrepository.com/artifact/org.apache.logging.log4j/log4j-core -->
		<dependency>
			<groupId>org.apache.logging.log4j</groupId>
			<artifactId>log4j-core</artifactId>
			<!--version>2.6.1</version -->
		</dependency>
	</dependencies>

	
	
#RUN on 172.31.28.230 of MDI Real Time OSS
/software/java/jdk1.7.0_79//bin/java -Dlog4j.configurationFile=log4j2.xml -jar CrossCheckMetadata-1.0.3-jar-with-dependencies.jar