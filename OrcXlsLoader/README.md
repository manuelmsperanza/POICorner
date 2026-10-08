# OrcXlsLoader

Launch com.hoffnungland.poi.corner.orcxlsloader.App with ConnectionName ExcelName SourcePath. Connection settings are loaded from etc/connections/<ConnectionName>.properties. SourcePath is the input directory prefix and should end with a path separator. The loader imports spreadsheet records into Oracle; use a test database when validating imports.

Build and test from the repository root:

~~~sh
mvn -B -pl OrcXlsLoader -am clean verify
~~~

See the [root README](../README.md) for requirements, dependencies, Javadoc, and release preparation.
