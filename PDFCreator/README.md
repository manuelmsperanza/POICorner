# PDFCreator

App creates Test.pdf in the current directory: a tagged PDF 2.0 document containing Hello world! The createPdf API accepts an output filename. RemovePdfPassword accepts source filename, destination filename, and password in that order, and rewrites the document without encryption.

Build and test from the repository root:

~~~sh
mvn -B -pl PDFCreator -am clean verify
~~~

See the [root README](../README.md) for requirements, dependencies, Javadoc, and release preparation.
