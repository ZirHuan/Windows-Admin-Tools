# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.1.0] - 2024-12-11

### Added
- Loop functionality: Press ENTER to run again or ESC to exit
- Screen clearing between runs for better user experience
- Improved exit message

### Changed
- Script now runs continuously until ESC is pressed
- Enhanced user prompts and instructions

## [1.0.0] - 2024-12-11

### Added
- Initial release
- Interactive numbered list of AD groups
- Group member display functionality
- Color-coded output
- Input validation
- Support for quitting with 'Q'
- Display of member object types (user, computer, group)
- Total member count display
- Handling of empty groups

### Fixed
- Replaced problematic ReadKey implementation for better compatibility