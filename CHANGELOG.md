# Changelog

All notable changes to this project will be documented in this file.

This project follows a simplified semantic versioning format:

MAJOR.MINOR.PATCH

---

## [1.1.0] - 2026-02-15

### Changed

- Calendar alignment updated from Sunday-based to Monday-based weeks.
- Adjusted date-to-column offset logic in `createSheet()`.

### Added

- Static unit tests for:
    - Monday-based header validation
    - Date-to-cell mapping
    - Template week rollover logic
    - Month boundary correctness
    - Week count calculation per month

### Technical Notes

- Introduced deterministic cell mapping validation to prevent regressions.
- Refactored `addTemplate()` to accept workbook parameter for testability.

---

## [1.0.0] - Initial Release

### Features

- 14-week template rotation
- Automatic month sheet creation
- Multi-shift support (D1, optional D2, N)
- Auto-formatting and page layout configuration
- CLI-based execution
