# Contributing to outlook-cli

Thanks for your interest in contributing!

## Getting Started

```sh
git clone https://github.com/yusufaltunbicak/outlook-cli.git
cd outlook-cli
pip install -e .
playwright install chromium
```

## Development

- Python 3.10+
- No test suite yet — contributions welcome
- Follow existing code style (type hints, docstrings where non-obvious)

## Submitting Changes

1. Fork the repo and create a feature branch
2. Make your changes
3. Test manually with `outlook` commands
4. Submit a pull request with a clear description

## Reporting Issues

- Use [GitHub Issues](https://github.com/yusufaltunbicak/outlook-cli/issues)
- Include your Python version, OS, and steps to reproduce

## Code Style

- Type hints for function signatures
- Click for CLI commands
- Rich for terminal output (stderr), stdout reserved for JSON
- httpx for HTTP requests

## Notes

- **Do not commit tokens, credentials, or personal data**
- Category management uses undocumented OWA endpoints — changes here need extra care
- All send/delete operations should include confirmation prompts by default
