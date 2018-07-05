#Create a new project
mvn archetype:generate -DarchetypeCatalog=http://repo.maven.apache.org/maven2/archetype-catalog.xml -Dfilter=maven-archetype-quickstart -DgroupId=me.hoffnungland -DartifactId=OrcXlsReport -Dpackage=me.hoffnungland.poi.corner.orcxlsreport -Dversion=0.0.1-SNAPSHOT
#Build settings
##Remove junit:junit:3.8.1