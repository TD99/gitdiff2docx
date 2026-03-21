# GitDiff2Docx Architecture

## System Architecture

```
┌─────────────────────────────────────────────────────────────┐
│                   GitDiff2Docx Application                  │
├─────────────────────────────────────────────────────────────┤
│                                                             │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────┐    │
│  │   CLI Mode   │  │ Interactive  │  │   GUI Mode   │    │
│  │ (Non-Inter.) │  │     Mode     │  │   (Tkinter)  │    │
│  └──────┬───────┘  └──────┬───────┘  └──────┬───────┘    │
│         │                 │                 │             │
│         └─────────────────┴─────────────────┘             │
│                           │                                │
│                    ┌──────▼──────┐                         │
│                    │  main.py    │                         │
│                    │ (Click CLI) │                         │
│                    └──────┬──────┘                         │
│                           │                                │
├───────────────────────────┼────────────────────────────────┤
│                           │                                │
│  ┌────────────────────────┼────────────────────────────┐  │
│  │         Core Package (gitdiff2docx/)               │  │
│  │                        │                            │  │
│  │  ┌──────┬──────┬──────┼──────┬──────┬──────┬────┐ │  │
│  │  │ cli  │config│ diff │ doc  │ git  │theme │util│ │  │
│  │  └──┬───┴───┬──┴───┬──┴───┬──┴───┬──┴───┬──┴──┬─┘ │  │
│  │     │       │      │      │      │      │     │   │  │
│  └─────┼───────┼──────┼──────┼──────┼──────┼─────┼───┘  │
│        │       │      │      │      │      │     │      │
│        │       │      │      │      │      │     │      │
├────────┼───────┼──────┼──────┼──────┼──────┼─────┼──────┤
│ Module Details                                            │
├───────────────────────────────────────────────────────────┤
│                                                             │
│  cli/                  config/              diff/          │
│  ├─ interactive.py     ├─ constants.py      └─ processor.py│
│  └─ __init__.py        ├─ loader.py                       │
│                        ├─ schema.py                        │
│                        └─ __init__.py                      │
│                                                             │
│  document/             git/                theme/          │
│  ├─ builder.py         ├─ operations.py    ├─ creator.py  │
│  ├─ layout.py          ├─ filtering.py     ├─ manager.py  │
│  ├─ media.py           └─ __init__.py      ├─ styling.py  │
│  ├─ tables.py                              └─ __init__.py │
│  └─ __init__.py                                           │
│                                                             │
│  utils/                localization/                       │
│  ├─ color.py           └─ loader.py                       │
│  ├─ console.py                                             │
│  ├─ dict.py                                                │
│  ├─ file.py                                                │
│  └─ __init__.py                                            │
│                                                             │
└─────────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────────┐
│                    External Resources                       │
├─────────────────────────────────────────────────────────────┤
│                                                             │
│  config.json      themes/*.json      lang/*.json           │
│  schemas/*.json   .gddignore                               │
│                                                             │
└─────────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────────┐
│                    External Systems                         │
├─────────────────────────────────────────────────────────────┤
│                                                             │
│  Git Repository  →  GitDiff2Docx  →  Word Document (.docx) │
│                                                             │
└─────────────────────────────────────────────────────────────┘
```

## Data Flow

```
User Input (CLI/GUI/Interactive)
        │
        ▼
┌───────────────┐
│  Parse Args   │
│  Load Config  │
└───────┬───────┘
        │
        ▼
┌───────────────┐
│  Load Theme   │
│  Load Lang    │
└───────┬───────┘
        │
        ▼
┌───────────────┐
│  Git Ops      │
│  - Get Files  │
│  - Filter     │
└───────┬───────┘
        │
        ▼
┌───────────────┐
│  Create Doc   │
│  - Header     │
│  - Legend     │
└───────┬───────┘
        │
        ▼
┌───────────────┐
│  Process File │  ◄── Loop for each changed file
│  - Get Bytes  │
│  - Calc Diff  │
│  - Add Table  │
└───────┬───────┘
        │
        ▼
┌───────────────┐
│  Save DOCX    │
│  Output       │
└───────────────┘
```

## Module Dependencies

```
main.py
  ├─→ config/
  │     ├─→ constants
  │     ├─→ loader
  │     └─→ schema
  │
  ├─→ theme/
  │     ├─→ creator
  │     ├─→ manager
  │     └─→ styling
  │
  ├─→ localization/
  │     └─→ loader
  │
  ├─→ git/
  │     ├─→ operations
  │     └─→ filtering
  │
  ├─→ document/
  │     ├─→ builder
  │     ├─→ layout
  │     ├─→ tables
  │     └─→ media
  │
  ├─→ diff/
  │     └─→ processor
  │
  ├─→ cli/
  │     └─→ interactive
  │
  └─→ utils/
        ├─→ console
        ├─→ color
        ├─→ file
        └─→ dict
```

## Key Design Patterns

### 1. Separation of Concerns
- Each module has a single responsibility
- Clear boundaries between layers
- Easy to test and maintain

### 2. Dependency Injection
- Configuration passed to modules
- Theme and language objects injected
- Allows easy testing and mocking

### 3. Facade Pattern
- `main.py` provides simple interface
- Hides complex subsystem interactions
- User-friendly CLI and GUI

### 4. Strategy Pattern
- Multiple output strategies (CLI, GUI)
- Different theme strategies
- Pluggable components

### 5. Factory Pattern
- Document creation in `builder.py`
- Theme loading in `manager.py`
- Configuration assembly

## Technology Stack

```
┌──────────────────────────────────┐
│     User Interface Layer         │
├──────────────────────────────────┤
│ Click (CLI)                      │
│ Tkinter (GUI)                    │
└──────────────────────────────────┘
           │
           ▼
┌──────────────────────────────────┐
│     Application Layer            │
├──────────────────────────────────┤
│ Python 3.7+                      │
│ GitDiff2Docx Package             │
└──────────────────────────────────┘
           │
           ▼
┌──────────────────────────────────┐
│     Library Layer                │
├──────────────────────────────────┤
│ python-docx (Word documents)     │
│ Pygments (Syntax highlighting)   │
│ Pillow (Image processing)        │
│ pathspec (File filtering)        │
└──────────────────────────────────┘
           │
           ▼
┌──────────────────────────────────┐
│     System Layer                 │
├──────────────────────────────────┤
│ Git (Version control)            │
│ OS (File system)                 │
│ Platform (Cross-platform)        │
└──────────────────────────────────┘
```

## Package Structure Details

```
gitdiff2docx/
│
├── __init__.py              # Package initialization
├── main.py                  # Entry point with Click CLI
│
├── cli/                     # Command-line interface
│   ├── __init__.py
│   └── interactive.py       # Interactive prompts
│
├── config/                  # Configuration handling
│   ├── __init__.py
│   ├── constants.py         # Global constants
│   ├── loader.py            # JSON loading & caching
│   └── schema.py            # JSON schema processing
│
├── diff/                    # Diff calculation
│   ├── __init__.py
│   └── processor.py         # Diff processing logic
│
├── document/                # Document generation
│   ├── __init__.py
│   ├── builder.py           # Document creation
│   ├── layout.py            # Table layouts & borders
│   ├── media.py             # Image handling
│   └── tables.py            # Diff & legend tables
│
├── git/                     # Git operations
│   ├── __init__.py
│   ├── operations.py        # Git commands
│   └── filtering.py         # .gddignore handling
│
├── localization/            # Multi-language
│   ├── __init__.py
│   └── loader.py            # Language file loading
│
├── theme/                   # Theme management
│   ├── __init__.py
│   ├── creator.py           # Theme creation
│   ├── manager.py           # Theme loading
│   └── styling.py           # Theme application
│
└── utils/                   # Utility functions
    ├── __init__.py
    ├── color.py             # Color conversion
    ├── console.py           # Terminal I/O
    ├── dict.py              # Dictionary merging
    └── file.py              # File detection
```

## Extension Points

The architecture supports easy extension:

1. **New Themes**: Add JSON files to `themes/`
2. **New Languages**: Add JSON files to `lang/`
3. **New Output Formats**: Extend `document/` module
4. **New CLI Commands**: Add to `main.py` Click commands
5. **New GUI Features**: Extend `gitdiff2docx-gui.py`

## Performance Considerations

- **Caching**: JSON schema files cached on first load
- **Lazy Loading**: Modules imported only when needed
- **Streaming**: Large files processed incrementally
- **Threading**: GUI uses separate thread for processing
- **Memory**: Binary files detected early to avoid loading

## Security Features

- **Path Validation**: Prevents directory traversal attacks
- **Input Sanitization**: Theme names validated
- **Safe File Operations**: Error handling for all I/O
- **No Code Execution**: Only data files processed
- **Cross-Platform**: Safe file opening per OS
