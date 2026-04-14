# Beamer Theme Reference

Quick reference for picking or overriding the `--theme` flag in `pptx_to_beamer.py`.
Pass any name below to `--theme` (case-sensitive).

## Automatically inferred mappings

The converter maps pptx color palettes to themes:

| Condition detected in .pptx               | Inferred theme  |
|-------------------------------------------|-----------------|
| pptx theme name contains "metropolitan"   | Metropolis      |
| pptx theme name contains "circuit"        | Berlin          |
| pptx theme name contains "facet"          | AnnArbor        |
| pptx theme name contains "ion"            | Warsaw          |
| Dark background (brightness < 60)         | Metropolis      |
| Blue-dominant background                  | Warsaw          |
| Red/maroon background                     | AnnArbor        |
| Green-dominant background                 | Hannover        |
| Fallback / unrecognised                   | **Madrid**      |

Override at any time:
```bash
python pptx_to_beamer.py slides.pptx --theme Metropolis
```

## Common Beamer themes

### Navigation-bar themes
| Theme        | Character                                |
|--------------|------------------------------------------|
| `Madrid`     | Blue headline, navigation dots at bottom |
| `Warsaw`     | Darker blue, section bar                 |
| `Berlin`     | Split header/footer, color blocks        |
| `AnnArbor`   | Tan/maroon, university-style             |
| `CambridgeUS`| Red and grey, clean                      |
| `Boadilla`   | Minimalist sidebar                       |

### Sidebar themes
| Theme        | Character                                |
|--------------|------------------------------------------|
| `Hannover`   | Left sidebar with section tree           |
| `Goettingen` | Right sidebar                            |
| `Marburg`    | Narrow left sidebar                      |
| `PaloAlto`   | Wide left sidebar                        |

### Minimal / modern themes
| Theme        | Character                                |
|--------------|------------------------------------------|
| `Metropolis` | Clean, dark or light, widely used today  |
| `default`    | No decoration, bare Beamer               |
| `plain`      | No header/footer at all                  |
| `Bergen`     | Rounded blocks, light                    |
| `Luebeck`    | Top navigation, compact                  |

## Color themes (apply on top of a structural theme)

You can combine a structural theme with a color theme by adding to the preamble:

```latex
\usecolortheme{dolphin}   % blue/grey
\usecolortheme{seahorse}  % blue tones
\usecolortheme{rose}      % red tones
\usecolortheme{whale}     % deep blue
\usecolortheme{crane}     % orange
\usecolortheme{beaver}    % dark red/grey
```

## Gotchas

- `Metropolis` requires the **Fira Sans** font and the `metropolis` package
  (`tlmgr install beamertheme-metropolis`). If Fira Sans is missing, fallback:
  ```latex
  \usefonttheme{default}
  ```
- Themes with navigation bars (`Madrid`, `Warsaw`) shrink your content area
  slightly — if content overflows a frame, add `[allowframebreaks]`:
  ```latex
  \begin{frame}[allowframebreaks]{Long slide title}
  ```
- The `plain` option removes all decorations from a single frame:
  ```latex
  \begin{frame}[plain]
    \titlepage
  \end{frame}
  ```
