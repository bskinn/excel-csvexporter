## CHANGELOG: CSV Exporter Excel VBA Add-In

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](http://keepachangelog.com/en/1.0.0/)
and this project adheres to [Semantic Versioning](http://semver.org/spec/v2.0.0.html).


### [Unreleased]

#### Changed
 * UserForm now reappears in its prior location when closed
   and re-opened, instead of always reappearing in the center
   of the Excel window.

### [1.0.0] - 2016-01-30

*Initial release*

#### Features
 * Folder selection works
 * Name, number format, and separator entry work
 * Append vs overwrite works
 * Modeless form retains folder/filename/format/separator/etc. within a given Excel instance

#### Limitations
 * Exports only a single contiguous range at a time

#### Internals
 * Modest validity checking implemented for filename
   * Red text and disabled `Export` button on invalid filename
 * No validity checking implemented for number format
 * Disabled `Export` button if number format or separator are empty
