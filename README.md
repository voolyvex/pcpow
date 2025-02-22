# PCPow

Simple Windows power management tool that safely closes apps before sleep/restart/shutdown.

## Usage

```bash
pcpow sleep     # Put PC to sleep
pcpow restart   # Restart PC
pcpow shutdown  # Shutdown PC
pcpow -h        # Show help
```

Add `-Force` to skip confirmation prompts:
```bash
pcpow sleep -Force
```

## Features

- Safely closes applications before power actions
- Confirms before closing apps (unless using -Force)
- Works from any directory
- Simple command-line interface

## Installation

1. Clone this repository
2. Add the directory to your system PATH
3. Run `pcpow -h` to verify installation

## License

MIT License - Feel free to use and modify 