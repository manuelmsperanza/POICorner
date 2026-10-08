# Maintenance notes

## Stable dependencies and build

Dependencies were reviewed on 2026-10-08. The parent now pins stable Maven 3
plugins, compiles with --release 25, and checks the Maven/JDK prerequisites.
POI uses matching 5.5.1 full OOXML schemas. Obsolete schema, contrib and security
artifacts were removed; contrib previously introduced Log4j 1 transitively.
See README.md for dependency versions and the maintenance commands.

The published oracleconn 23.26.3.0.0.55 POM contains invalid legacy Oracle XDK
system paths. Maven reports that upstream warning and skips its transitive
dependencies. Both Oracle modules explicitly declare ojdbc8 and ucp 23.26.3.0.0;
the shared dbconn dependency comes from XlsReport. Fixing the published wrapper
POM requires a separate release in its owning repository.

## Documentation and tests

Root and module READMEs now describe build requirements, APIs, CLI arguments,
input directories, and release behavior. Javadoc for JSON export, XML headers,
and PDF creation was expanded. Generated API documentation is checked during
release preparation.

Placeholder tests were replaced with checks for JSON file filtering, nested
JSON XLSX output, worksheet name boundaries, XML template header mapping, Visio
text normalization, tagged PDF creation, and incomplete database CLI arguments.
An in-memory H2 test verifies JDBC text, number, and NULL export without a server.
These checks do not cover live Oracle/PostgreSQL connections or full production
spreadsheet import workflows.
