# XmlXlsReport

Template-based XML-to-Excel support. XmlToXlsManager loads a workbook template, indexes named headers with NodeSheet, and maps XML data into worksheets. Template header rows must exist and contain string-valued field names. The App entry point is a greeting demo; use the manager API for XML exports.

Build and test from the repository root:

~~~sh
mvn -B -pl XmlXlsReport -am clean verify
~~~

See the [root README](../README.md) for requirements, dependencies, Javadoc, and release preparation.
