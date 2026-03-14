# Design System — Dark Professional Desktop UI

Extracted from the Spec Critic application GUI. This document defines the complete visual language: color palette, typography, spacing, component patterns, animation system, and layout architecture. Built on CustomTkinter (dark mode), but the design tokens and patterns are framework-agnostic.

---

## Philosophy

The UI follows a **dark, card-based layout** with minimal chrome. The aesthetic is closer to a modern IDE or dashboard than a typical desktop app — muted backgrounds, high-contrast accent colors on interactive elements, monospace text for data, and generous padding. The overall feel is professional and dense without being cluttered.

Key principles:

- **Cards on dark canvas** — content lives in slightly elevated `bg_card` panels on a near-black `bg_dark` background. No borders between sections; elevation (color difference) creates hierarchy.
- **Color as information** — severity levels, confidence tiers, and verification verdicts each have a dedicated color. Color is never decorative.
- **Minimal labels, maximal density** — section headers are small, uppercase, muted-gray labels (`INPUTS`, `FILES`, `ACTIVITY LOG`, `SUMMARY`). Content does the talking.
- **Progressive disclosure** — cards and panels are collapsible. Token gauge, file list, activity log, and individual finding cards all expand/collapse with a single click.
- **Monospace for data, sans-serif for UI** — Consolas for tokens, file paths, counts, and log entries. Segoe UI for labels, headings, and body text.

---

## Color Palette

### Core Backgrounds

| Token             | Hex       | Usage                                      |
|-------------------|-----------|--------------------------------------------|
| `bg_dark`         | `#0D0D0D` | App window background (the canvas)         |
| `bg_card`         | `#1A1A1A` | Card/panel surfaces                        |
| `bg_input`        | `#252525` | Input fields, text areas, nested containers|
| `border`          | `#333333` | Borders, dividers, subtle outlines         |

### Text Hierarchy

| Token             | Hex       | Usage                                        |
|-------------------|-----------|----------------------------------------------|
| `text_primary`    | `#FFFFFF` | Headings, important values, filenames        |
| `text_secondary`  | `#B0B0B0` | Body text, descriptions, labels              |
| `text_muted`      | `#707070` | Hints, metadata, disabled text, section tags |

### Accent & Interactive

| Token             | Hex       | Usage                                      |
|-------------------|-----------|--------------------------------------------|
| `accent`          | `#3B82F6` | Primary buttons, progress bars, links      |
| `accent_hover`    | `#2563EB` | Button hover state                         |
| `accent_glow`     | `#60A5FA` | Animated glow on completion                |

### Status Colors

| Token             | Hex       | Usage                               |
|-------------------|-----------|--------------------------------------|
| `success`         | `#22C55E` | Completion, confirmed verdicts       |
| `success_glow`    | `#4ADE80` | Animated success pulse               |
| `warning`         | `#F59E0B` | Warnings, corrected verdicts, amber  |
| `error`           | `#EF4444` | Errors, disputed verdicts            |

### Severity Scale

Used for classification badges, card accent strips, and grouped section headers.

| Token      | Hex       | Context                              |
|------------|-----------|--------------------------------------|
| `critical` | `#DC2626` | Highest severity — showstoppers      |
| `high`     | `#F97316` | Significant errors                   |
| `medium`   | `#EAB308` | Reference errors, inconsistencies    |
| `gripe`    | `#A855F7` | Cosmetic / editorial                 |

### Verification Verdicts

| Verdict      | Hex       | Icon |
|--------------|-----------|------|
| `CONFIRMED`  | `#22C55E` | ✓    |
| `CORRECTED`  | `#F59E0B` | ✎    |
| `UNVERIFIED` | `#6B7280` | —    |
| `DISPUTED`   | `#EF4444` | ✗    |

### Confidence Tiers

| Tier     | Range     | Hex       |
|----------|-----------|-----------|
| High     | ≥ 85%     | `#22C55E` |
| Moderate | 60%–84%   | `#F59E0B` |
| Low      | < 60%     | `#EF4444` |

### Special

| Token          | Hex       | Usage                              |
|----------------|-----------|------------------------------------|
| `coordination` | `#06B6D4` | Cross-spec / coordination section  |

---

## Typography

### Font Stack

| Role       | Font Family | Fallback     | Usage                                           |
|------------|-------------|--------------|--------------------------------------------------|
| UI         | Segoe UI    | system-ui    | All labels, headings, body text, buttons         |
| Data       | Consolas    | monospace    | Token counts, file paths, log entries, API keys  |

### Type Scale

| Element                    | Font   | Size | Weight | Color            |
|----------------------------|--------|------|--------|------------------|
| App title                  | Segoe  | 28   | Bold   | `text_primary`   |
| App subtitle               | Segoe  | 13   | Normal | `text_secondary` |
| Card section label         | Segoe  | 11   | Bold   | `text_muted`     |
| Input label                | Segoe  | 12   | Normal | `text_secondary` |
| Input value                | Consolas| 12  | Normal | `text_primary`   |
| Button text (primary)      | Segoe  | 14   | Bold   | `text_primary`   |
| Button text (secondary)    | Segoe  | 11   | Normal | `text_secondary` |
| Segmented button labels    | Segoe  | 11   | Normal | `text_secondary` |
| Hint / descriptor text     | Segoe  | 10   | Normal | `text_muted`     |
| File list filename         | Segoe  | 11   | Normal | `text_secondary` |
| File list token count      | Consolas| 10  | Normal | `text_muted`     |
| Log entry text             | Consolas| 12  | Normal | varies by level  |
| Report heading             | Segoe  | 18   | Bold   | `text_primary`   |
| Summary stat number        | Segoe  | 22   | Bold   | severity color   |
| Summary stat label         | Segoe  | 9    | Bold   | `text_muted`     |
| Finding card filename      | Segoe  | 12   | Bold   | `text_primary`   |
| Finding card section ref   | Segoe  | 11   | Normal | `text_muted`     |
| Finding body text          | Segoe  | 12   | Normal | `text_secondary` |
| Existing text (in finding) | Consolas| 11  | Normal | `error` red      |
| Replacement text           | Consolas| 11  | Normal | `success` green  |
| Code reference             | Segoe  | 11   | Normal | `accent` blue    |

---

## Spacing & Layout

### Global

| Property          | Value    |
|-------------------|----------|
| Window padding    | 24px all sides |
| Card corner radius| 8px      |
| Input corner radius| 8px (fields), 4px (nested containers) |
| Card internal padding (horizontal) | 16px |
| Card internal padding (vertical)   | 12px |
| Section gap (between cards)        | 12–16px |

### Input Grid

The INPUTS card uses a 2-column grid layout:

- **Column 0**: Label (100px fixed width, left-aligned)
- **Column 1**: Control (expands to fill, 8px left margin from label)
- **Row padding**: 8px vertical between rows

### Component Heights

| Component           | Height |
|----------------------|--------|
| Text input           | 36px   |
| Browse button        | 36px   |
| Segmented button     | 32px   |
| Checkbox             | 20×20px|
| Primary action button| 44px   |
| Secondary button     | 30px   |
| Toolbar button       | 24–30px|
| Progress bar         | 4px    |

---

## Component Patterns

### Card (Collapsible Section)

Every major section is a **collapsible card** with a consistent structure:

```
┌─────────────────────────────────────────────┐
│  ▼  SECTION LABEL              count / info │  ← header (clickable, cursor: hand2)
│─────────────────────────────────────────────│
│                                             │  ← content container (toggled)
│    [content goes here]                      │
│                                             │
└─────────────────────────────────────────────┘
```

- **Container**: `bg_card`, corner radius 8
- **Header**: Transparent frame, padx 16, pady 12
- **Arrow**: Consolas 12, `text_muted`, `▼` when expanded / `▶` when collapsed
- **Section label**: Segoe UI 11 bold, `text_muted`, ALL CAPS
- **Right-aligned info**: Segoe UI 11, `text_secondary` (e.g., "4/6 selected")
- **Content**: Transparent frame, padx 16, pady (0, 16)
- The entire header row is click-bound for toggle

### Token Gauge (Capacity Bar)

A horizontal capacity indicator with animated fill:

- **Bar container**: `bg_input`, corner radius 4, height 8px
- **Fill bar**: Corner radius 4, height 8px, animated width
- **Color transitions**: `success` → `warning` → `error` based on fill percentage
  - 0–70%: `success` green
  - 70–90%: `warning` amber
  - 90–100%: `warning` amber with warning text
  - >100%: `error` red with "Capacity Exceeded!" message
- **Status text**: Below bar, Segoe UI 11
- **Count label**: Right-aligned in header, Consolas 12 (e.g., "42,380 / 150,000")

### Primary Action Button

The main action button has three visual states:

1. **Ready**: `accent` background, `accent_hover` on hover, "Run Review" label, Segoe UI 14 bold, 44px tall, corner radius 8
2. **Processing**: Disabled state, pulsing animation that oscillates between `bg_input` and `accent`, "Processing..." label
3. **Complete**: `success` green background with a brief glow animation (success → success_glow → success), "✓ Complete" label, disabled

### Secondary Buttons

Used for toolbar actions, browse, and utility functions:

- **Background**: `bg_input` (appears as slightly raised from card)
- **Border**: 1px `border`
- **Hover**: `border` background
- **Text**: `text_secondary`
- **Font**: Segoe UI 11 or 12
- **Height**: 30px (toolbar) or 36px (input row)
- **Corner radius**: 6px

### Segmented Button (Mode Selector)

Used for model selection, review mode, and output mode:

- **Selected state**: `accent` background, `accent_hover` on hover
- **Unselected state**: `bg_input` background, `border` on hover
- **Text**: `text_secondary` (unselected), inherited (selected)
- **Height**: 32px
- **Font**: Segoe UI 11
- **Hint text**: 10px `text_muted`, 12px left margin after the segmented button

### Checkbox

- **Size**: 20×20px
- **Checked fill**: `accent`
- **Hover**: `accent_hover`
- **Border**: `border`
- **Checkmark**: `text_primary` (white)
- **Corner radius**: 4px
- **Border width**: 2px

### Text Input

- **Background**: `bg_input`
- **Border**: `border` (2px for textbox, default for entry)
- **Text**: `text_primary`, Consolas 12
- **Placeholder text**: `text_muted`
- **Height**: 36px (single-line), 80px (multiline/textbox)
- **Password masking**: Bullet character `•`

### Progress Bar

- **Height**: 4px
- **Corner radius**: 2px
- **Track**: `bg_input`
- **Fill**: `accent`
- Appears between the action button and the log, packed with `fill="x"`

---

## Activity Log

The log is a scrollable text widget inside a collapsible card:

- **Container**: `bg_input`, corner radius 4, Consolas 12
- **Timestamp format**: `[HH:MM:SS]  message` (9-space indent for continuation lines)
- **Line colors by level**:

| Level   | Color            | Prefix |
|---------|------------------|--------|
| info    | `text_secondary` | (none) |
| success | `success`        | ✓      |
| warning | `warning`        | ⚠      |
| error   | `error`          | ✗      |
| step    | `accent`         | ▸      |
| file    | `text_primary`   | →      |
| muted   | `text_muted`     | (none) |

- **Paced output**: Log entries are queued and rendered with delays to create a "live" feel:
  - File entries: 200ms delay
  - Status entries: 400ms delay
  - Non-paced entries render immediately

---

## Finding Cards

Each finding is a collapsible card with a colored severity accent strip on the left:

```
┌────┬──────────────────────────────────────────────┐
│    │  ▼ [CRITICAL] 95%  — filename.docx • Part 2  │  ← header
│ ██ │──────────────────────────────────────────────│
│ ██ │  Issue description text...                    │  ← body
│ ██ │                                               │
│ ██ │  Existing:     old text in red                │
│ ██ │  Replace with: new text in green              │
│ ██ │  Reference:    code ref in blue               │
│ ██ │                                               │
│ ██ │  ✓ CONFIRMED   Explanation text...            │  ← verification
└────┴──────────────────────────────────────────────┘
```

### Structure

- **Outer frame**: `fg_color` = severity color, corner radius 8 (creates the accent strip)
- **Inner card**: `bg_input`, corner radius 6, padx (4, 0) — the 4px left gap reveals the colored outer frame as the accent strip
- **Header**: Clickable, contains arrow + severity badge + confidence + filename + section
- **Body**: Toggleable, padx 14, pady (4, 12)

### Severity Badge

- Font: Segoe UI 10 bold
- Background: severity color
- Text: white (except MEDIUM uses black for contrast on yellow)
- Width: 70px, height: 22px, corner radius 4

### Confidence Badge

- Font: Consolas 9 bold
- No background — text only
- Color: confidence tier color (green/amber/red)
- Width: 36px

### Verification Row

- Verdict badge: Segoe UI 10 bold, `fg_color` = verdict color, white text (except CORRECTED uses black), 100px wide, 22px tall, corner radius 4
- Explanation: Segoe UI 11, `text_secondary`, wraplength 550
- Correction text: Consolas 11, `warning` amber
- Sources: Consolas 10, `text_muted`, max 3 URLs shown

### Data Row Pattern (within finding body)

Used for Existing/Replace/Reference/Correction rows:

```
Label:          Value text
```

- Label: Segoe UI 11 bold, `text_muted`, 90px fixed width, left-aligned
- Value: Consolas 11 (for text content) or Segoe UI 11 (for references)
- Color-coded: existing=`error`, replacement=`success`, reference=`accent`, correction=`warning`
- Row spacing: 3pt after each row (via `space_after`)

---

## Summary Grid

The report header uses a grid of stat cells:

```
┌──────────┬──────────┬──────────┬──────────┬──────────┐
│    4     │    7     │   12     │    3     │   26     │
│ CRITICAL │   HIGH   │  MEDIUM  │  GRIPES  │  TOTAL   │
└──────────┴──────────┴──────────┴──────────┴──────────┘
```

- **Cell container**: `bg_input`, corner radius 6, padx 4 between cells
- **Number**: Segoe UI 22 bold, colored by severity
- **Label**: Segoe UI 9 bold, `text_muted`, ALL CAPS
- **Internal padding**: padx 12, pady 10
- **Grid**: Equal column weights, cells expand to fill
- **Optional extra column**: "CROSS-CHECK" with `coordination` cyan color

---

## Animation System

### Constants

| Token              | Value  | Usage                                    |
|--------------------|--------|------------------------------------------|
| `log_file_delay`   | 200ms  | Delay between file log entries           |
| `log_status_delay` | 400ms  | Delay between status log entries         |
| `gauge_step`       | 33ms   | Frame interval for gauge animation       |
| `gauge_duration`   | 700ms  | Total gauge fill animation time          |
| `fade_duration`    | 200ms  | Fade in/out transitions                  |
| `fade_steps`       | 8      | Number of frames in a fade               |
| `pulse_interval`   | 1500ms | Full cycle of processing pulse           |
| `pulse_step_ms`    | 67ms   | Frame interval for pulse animation       |
| `glow_step_ms`     | 67ms   | Frame interval for glow animations       |
| `expand_duration`  | 200ms  | Expand/collapse panel animation          |
| `expand_steps`     | 10     | Number of frames in expand/collapse      |

### Easing

- **Gauge fill**: `ease_out_cubic` — `1 - pow(1 - t, 3)` — fast start, gentle deceleration
- **Color blending**: Linear interpolation between hex colors (per-channel RGB lerp)
- **Processing pulse**: Sinusoidal oscillation — `(sin(step / steps * π * 2) + 1) / 2` — smooth loop between `bg_input` and `accent`
- **Completion glow**: Quick flash from `success` → `success_glow` → `success` over 15 frames
- **Error glow**: Sinusoidal pulse on the file panel title between `error` and `#ff9999`

### Utility Functions

```python
def lerp(start, end, t):
    return start + (end - start) * t

def ease_out_cubic(t):
    return 1 - pow(1 - t, 3)

def hex_to_rgb(h):
    h = h.lstrip('#')
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))

def rgb_to_hex(r, g, b):
    return f"#{r:02x}{g:02x}{b:02x}"

def blend_colors(c1, c2, t):
    r1, g1, b1 = hex_to_rgb(c1)
    r2, g2, b2 = hex_to_rgb(c2)
    return rgb_to_hex(
        int(lerp(r1, r2, t)),
        int(lerp(g1, g2, t)),
        int(lerp(b1, b2, t))
    )
```

---

## Window & Layout Structure

### Main Window

- **Default size**: 900×950
- **Minimum size**: 750×700
- **Background**: `bg_dark`
- **Appearance mode**: Dark
- **Color theme**: Blue (CustomTkinter default)

### Layout Stack (top to bottom)

All content lives in a single container frame (`fg_color="transparent"`, padx 24, pady 24):

1. **Header** — App title + subtitle + utility buttons (right-aligned)
2. **INPUTS card** — Collapsible, contains all input controls in a grid
3. **FILES panel** — Collapsible, appears after file selection
4. **TOKEN CAPACITY gauge** — Collapsible, animated fill bar
5. **Primary action button** — Full-width
6. **Progress bar** — 4px, appears during processing (packed dynamically)
7. **ACTIVITY LOG** — Collapsible, expands to fill remaining vertical space

The log is the only element with `expand=True` — it absorbs all remaining vertical space.

### Report Window (Pop-out)

A separate top-level window with its own toolbar:

- **Size**: 960×800, min 700×500
- **Background**: `bg_dark`
- **Toolbar**: `bg_card`, 48px fixed height, contains title + export buttons
- **Body**: Scrollable frame, transparent, padx 16, pady 16
- **Content stack**: Summary grid → Alerts → Findings → Cross-check → Reviewer's notes

---

## Dialog Pattern

Modal dialogs (e.g., "How It Works", "Pending Batch Found") follow this structure:

- **Overlay**: `CTkToplevel`, transient to main window, grab_set for modality
- **Background**: `bg_dark`
- **Content panel**: `bg_card`, corner radius 8, padx/pady 16
- **Title**: Segoe UI 16–20 bold, `text_primary`
- **Subtitle/info**: Segoe UI 12 or Consolas 11, `text_muted` or `text_secondary`
- **Action buttons**: Bottom of panel, left-aligned
  - Primary: `accent` background, `accent_hover` hover, Segoe UI 12, corner radius 6
  - Secondary: `bg_input` background, 1px `border`, `text_secondary`
- **Close**: Either a dedicated "Close" button or dialog destroy on action

---

## Color Application Rules

1. **Never use color for decoration alone.** Every color carries semantic meaning.
2. **Severity colors appear in three places**: finding card accent strip, severity badge background, and section header text.
3. **Text on colored backgrounds**: Use white text on all severity colors except MEDIUM (use black on yellow for contrast).
4. **Nested elevation**: `bg_dark` → `bg_card` → `bg_input`. Never skip a level. An input field inside a card uses `bg_input`, not `bg_dark`.
5. **Interactive elements**: Default = `bg_input` with `border` outline. Hover = `border` fill. Active/selected = `accent`.
6. **Data values in findings**: existingText = red (`error`), replacementText = green (`success`), codeReference = blue (`accent`), correction = amber (`warning`).

---

## Adapting to Other Frameworks

If porting to React / Electron / web:

- `bg_dark` → `background-color` on `<body>` or root container
- `bg_card` → card component background
- `bg_input` → `input`, `textarea`, `select` backgrounds
- Font sizes are in **points** (CTk convention). Multiply by ~1.33 for CSS pixels (e.g., 11pt ≈ 15px, 12pt ≈ 16px, 14pt ≈ 19px)
- All corner radii are in pixels and transfer directly to CSS `border-radius`
- Animation timings transfer directly to CSS `transition-duration` or JS `requestAnimationFrame` intervals
- The color blending / easing functions work identically in any language

If porting to another Python GUI framework (PyQt, Kivy, etc.):

- The color tokens, type scale, and spacing values transfer directly
- Replace `CTkFrame` → equivalent container widget
- Replace `self.after(ms, callback)` → framework-specific timer
- The collapsible card pattern maps to any show/hide toggle on a container
