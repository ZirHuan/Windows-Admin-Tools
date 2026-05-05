# Contributing to AD Group User Listing Script

Thank you for considering contributing to this project! Here are some guidelines to help you get started.

## How Can I Contribute?

### Reporting Bugs

Before creating bug reports, please check existing issues to avoid duplicates. When creating a bug report, include:

- A clear and descriptive title
- Steps to reproduce the issue
- Expected behavior
- Actual behavior
- Your environment (PowerShell version, Windows version, AD environment)
- Any relevant logs or error messages

### Suggesting Enhancements

Enhancement suggestions are welcome! Please provide:

- A clear and descriptive title
- A detailed description of the proposed enhancement
- Examples of how the enhancement would be used
- Why this enhancement would be useful to most users

### Pull Requests

1. Fork the repository
2. Create a new branch from `main` (`git checkout -b feature/YourFeature`)
3. Make your changes
4. Test your changes thoroughly
5. Commit your changes with clear, descriptive messages
6. Push to your fork
7. Open a Pull Request with a clear description of the changes

## Code Style Guidelines

- Follow PowerShell best practices
- Use meaningful variable names
- Add comments for complex logic
- Use proper indentation (4 spaces)
- Include error handling where appropriate
- Keep functions focused and concise
- Use approved PowerShell verbs for function names

## Testing

Before submitting a PR:

- Test in PowerShell Console (not ISE due to ReadKey limitations)
- Test in both PowerShell 5.1 and PowerShell 7+ if possible
- Test with various AD group scenarios (empty groups, nested groups, etc.)
- Verify the loop functionality (ENTER and ESC keys)
- Test with different terminal sizes and color schemes

## Development Setup

1. Clone the repository
2. Ensure you have the Active Directory PowerShell module installed
3. Have access to a test AD environment (don't test on production!)

## Questions?

Feel free to open an issue with your question or reach out to the maintainers.

Thank you for your contributions!