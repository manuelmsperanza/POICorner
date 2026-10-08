# OrcXlsReport

Run the module's App class with:

~~~text
ConnectionName ProjectName ExcelName TargetPath [id] [name]
~~~

Connection settings are read from etc/connections/<ConnectionName>.properties.
SQL files are read from <ProjectName>/queries, with optional parameterized SQL
in queriesById and queriesByName. TargetPath is an output directory prefix;
include a trailing path separator. ExcelName is the filename without .xlsx.
Database access and SQL inputs are required for actual reports.

Build and test from the repository root:

~~~sh
mvn -B -pl OrcXlsReport -am clean verify
~~~

See the [root README](../README.md) for build requirements and release preparation.
