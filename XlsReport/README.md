# XlsReport

Shared ExcelManager library for JDBC ResultSet and JSON exports. Construct ExcelManager with a file prefix, call getQueryResult or getJsonResult, and call finalWrite once. finalWrite appends .xlsx and closes the workbook. Its targetPath is a directory prefix and must end with a path separator. Do not reuse the manager after writing. JSON scalar values are currently exported as text.

Build and test from the repository root:

~~~sh
mvn -B -pl XlsReport -am clean verify
~~~

See the [root README](../README.md) for requirements, dependencies, Javadoc, and release preparation.
