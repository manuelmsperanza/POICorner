#Create a new project
mvn archetype:generate -DarchetypeCatalog=http://repo.maven.apache.org/maven2/archetype-catalog.xml -Dfilter=maven-archetype-quickstart -DgroupId=net.dtdns.hoffunungland -DartifactId=ExtractShapesVsdx -Dpackage=net.dtdns.hoffunungland.poi.corner.extractshapesvsdx -Dversion=0.0.1-SNAPSHOT
#Build settings
##Remove junit:junit:3.8.1