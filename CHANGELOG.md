## CHANGELOG: CSV Exporter

#### v1.0.0: Initial release

*Features*
 * Folder selection works
 * Name, number format, and separator entry work
 * Append vs overwrite works
 * Modeless form retains folder/filename/format/separator/etc. within a given Excel instance

*Limitations*
 * Exports only a single contiguous range at a time

*Internals*
 * Modest validity checking implemented for filename
   * Red text and disabled `Export` button on invalid filename
 * No validity checking implemented for number format
 * Disabled `Export` button if number format or separator are empty
