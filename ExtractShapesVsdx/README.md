# ExtractShapesVsdx

Launch com.hoffnungland.poi.corner.extractshapesvsdx.App from a directory containing .vsdx files. The application scans that directory, reads each Visio page, and logs normalized shape text. File matching uses the lowercase .vsdx extension. Configure Log4j output as needed.

Build and test from the repository root:

~~~sh
mvn -B -pl ExtractShapesVsdx -am clean verify
~~~

See the [root README](../README.md) for requirements, dependencies, Javadoc, and release preparation.
