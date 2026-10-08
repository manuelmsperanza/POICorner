# JSONXlsReport

Swing application for converting selected JSON files into XLSX worksheets. Launch com.hoffnungland.poi.corner.jsonxlsloader.App with the module and dependency JARs on the classpath. Choose JSON files, an output directory, and a workbook name, then convert. A graphical desktop is required. Nested objects and arrays are expanded into worksheet columns.

Build and test from the repository root:

~~~sh
mvn -B -pl JSONXlsReport -am clean verify
~~~

See the [root README](../README.md) for requirements, dependencies, Javadoc, and release preparation.
