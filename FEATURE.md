# Feature Notes

## Build reliability
- Removed the parent enforcer plugin execution that attempted to resolve additional plugins during `mvn install` in constrained environments.
- Kept plugin management focused on compiler, release, and surefire defaults inherited by modules.

## Java 25 upgrade
- Updated the root Maven Java source/target properties to `25` so all modules inherit Java 25 compilation settings.

## Testing improvements
- Added focused unit coverage in `XlsReport` for:
  - the generated greeting message,
  - worksheet exception message propagation.

## Javadoc cleanup
- Reworked class and method Javadocs in `XlsReport` app and exception classes to use clear, valid tags and descriptions.
