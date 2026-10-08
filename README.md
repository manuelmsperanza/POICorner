# POICorner

Java utilities for exporting database and JSON content to Excel, loading spreadsheets into Oracle, extracting Visio text, and creating PDF documents.

## Requirements and build

Use JDK 25 or newer and Maven 3.6.3 or newer. Compilation targets Java 25.
The com.hoffnungland database and logging libraries must be available in your
local Maven repository or a repository configured in your Maven settings.
GitHub Packages credentials belong in ~/.m2/settings.xml, never in this repository.

Run from the repository root:

~~~sh
mvn -B clean verify
mvn -B javadoc:aggregate
~~~

Tests use JUnit Jupiter, temporary files, and an in-memory H2 database; they do not need live database
connections. API documentation is generated under target/reports/apidocs
(or target/site/apidocs, depending on the Maven report output settings).

## Modules

| Module | Purpose |
| --- | --- |
| XlsReport | Shared streaming Excel exporter for JDBC results and nested JSON |
| JSONXlsReport | Swing application for selecting JSON files and exporting XLSX |
| OrcXlsReport | Oracle query and metadata reports |
| H2XlsReport | H2 query reports |
| PgXlsReport | PostgreSQL query reports |
| OrcXlsLoader | Loads spreadsheet records into Oracle |
| XmlXlsReport | XML-to-Excel template support |
| ExtractShapesVsdx | Logs shape text from .vsdx files in the current directory |
| PDFCreator | Creates a tagged greeting PDF and rewrites password-protected PDFs |

See each module's README for usage. Database reports resolve connection settings
from etc/connections/<ConnectionName>.properties and SQL from
<ProjectName>/queries. Keep credentials and production inputs outside version control.

## Dependency maintenance

Dependency versions were reviewed on 2026-10-08 against Maven repository metadata.
Current versions include POI 5.5.1, iText 9.8.0, JUnit Jupiter 6.1.3, and
org.json 20260814. Stable private wrappers are log4j 2.26.1.41, dbconn 0.0.42,
oracleconn 23.26.3.0.0.55, h2dbconn 2.5.252.32, and pgdbconn 42.7.13.14;
these were checked against installed local release metadata.

Legacy POI schema artifacts are replaced with poi-ooxml-full 5.5.1; unused
contrib and security artifacts are removed, including their obsolete Log4j 1 dependency. Beta, milestone,
release-candidate and SNAPSHOT library upgrades are excluded. Reactor module
SNAPSHOT dependencies are intentional and are rewritten together by Maven Release.

~~~sh
mvn -B versions:display-dependency-updates versions:display-plugin-updates -DallowSnapshots=false
~~~

Review the result before editing: allowSnapshots=false alone does not exclude
beta or milestone releases. Stable Maven 3 plugins are pinned in the parent POM.
See the [Apache plugin catalog](https://maven.apache.org/plugins/) and
[POI release page](https://poi.apache.org/download.cgi).

## Release preparation

Commit verified source, documentation, and POM changes first. Use a clean Git
working tree, a configured Git author, and SSH access to the origin repository.

~~~sh
mvn -B release:prepare
~~~

Preparation runs clean verify javadoc:aggregate, removes SNAPSHOT suffixes,
commits the release POMs, tags the parent version, then commits the next
development versions. By default the release plugin pushes commits and the tag.
It does not publish packages; release:perform is a separate step.

If preparation fails, inspect release.properties and the console log, correct
the failure, and rerun release:prepare to resume. Do not delete release state
or roll back blindly after SCM commits or a tag have been created.
