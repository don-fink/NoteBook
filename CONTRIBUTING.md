# Contributing to NoteBook

Thank you for your interest in contributing! This document provides guidelines for contributing to NoteBook.

## Getting Started

1. Fork the repository on GitHub
2. Clone your fork locally
3. Set up the development environment (see [README.md](README.md#build-from-source))
4. Create a branch for your changes

## Development Setup

```powershell
# Create virtual environment
python -m venv .venv
& .\.venv\Scripts\Activate.ps1
python -m pip install -r requirements-dev.txt

# Run the app
python main.py
```

## Making Changes

### Code Style

- Use [Black](https://black.readthedocs.io/) for formatting (line length: 100)
- Use [isort](https://pycqa.github.io/isort/) for import sorting
- Follow PEP 8 conventions

Format your code before committing:

```powershell
black .
isort .
```

### Commit Messages

- Use clear, descriptive commit messages
- Start with a verb (Add, Fix, Update, Remove, etc.)
- Keep the first line under 72 characters

Examples:
- `Add keyboard shortcut for indent/outdent`
- `Fix crash when opening empty notebook`
- `Update README installation instructions`

## Pull Requests

1. Create a new branch from `main` for your feature or fix
2. Make your changes with clear commits
3. Test your changes thoroughly
4. Push to your fork
5. Open a Pull Request with a clear description

### PR Guidelines

- Keep changes focused — one feature or fix per PR
- Update documentation if needed
- Add yourself to contributors if this is your first contribution

## Reporting Issues

When reporting bugs, please include:

- Steps to reproduce the issue
- Expected vs. actual behavior
- Your Python version and OS
- Relevant log output (`crash.log` or `native_crash.log`)

## Feature Requests

Feature requests are welcome! Please:

- Check existing issues first to avoid duplicates
- Describe the use case and why it would be helpful
- Be open to discussion about implementation

## Questions?

Feel free to open an issue for questions about the codebase or contribution process.

---

Thank you for helping improve NoteBook!
