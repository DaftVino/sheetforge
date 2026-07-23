# Sheet Forge Design System

## 1. Core Color Palette

### Primary Navy

| Token | Hex | Use |
|---|---:|---|
| Navy 950 | `#071525` | Main screen background |
| Navy 900 | `#0A1B2E` | Primary app surface |
| Navy 850 | `#102338` | Cards and elevated panels |
| Navy 800 | `#172A3D` | Secondary cards and input backgrounds |
| Navy 700 | `#24384C` | Borders and inactive controls |
| Navy 600 | `#35495D` | Muted icons and supporting UI |

### Forge Red

| Token | Hex | Use |
|---|---:|---|
| Red 600 | `#C93636` | Dark accent states |
| Red 500 | `#DE4141` | Primary brand accent |
| Red 450 | `#E64A45` | Buttons, active toggles, and highlighted icons |
| Red 400 | `#ED5A53` | Hover or pressed emphasis |
| Red 200 | `#8F3134` | Muted red backgrounds |
| Red Tint | `#3B2028` | Badges and subtle accent surfaces |

### Neutral Grey

| Token | Hex | Use |
|---|---:|---|
| Grey 50 | `#F5F7F9` | Primary text |
| Grey 100 | `#E1E6EA` | Secondary light text |
| Grey 300 | `#AEB8C2` | Supporting text |
| Grey 400 | `#8794A1` | Metadata and inactive labels |
| Grey 500 | `#667482` | Disabled text |
| Grey 700 | `#334252` | Dividers and borders |
| Grey 800 | `#202F3F` | Subtle secondary surfaces |

## 2. Supporting Status Colors

| State | Hex | Use |
|---|---:|---|
| Success | `#43B875` | Active workbook status |
| Success Dark | `#183B32` | Installed or active pill background |
| Spreadsheet Green | `#3FA968` | Workbook and spreadsheet icons |
| Blue | `#566FAE` | Configuration and workflow icons |
| Purple | `#7855A8` | Reconciliation or sync modules |
| Warning | `#D49A48` | Attention states |
| Disabled | `#708091` | Off toggles and unavailable states |

---

# Visual Direction

## Brand Personality

Sheet Forge should feel:

- Technical without appearing developer-only
- Structured, reliable, and configurable
- Premium but practical
- Powerful without visual complexity
- More like a professional workflow platform than a spreadsheet utility

The design should avoid playful startup aesthetics, bright gradients, cartoon illustrations, and excessive glassmorphism.

---

# Logo System

## Primary Mark

The logo combines three visual concepts:

- An anvil or forged metal form
- Spreadsheet cells or table rows
- Modular blocks being assembled

### Logo Construction

- Main symbol width-to-height ratio: approximately `1.15:1`
- Use a simplified anvil silhouette on the left
- Use a three-column spreadsheet grid on the right
- Highlight two or three cells in Forge Red
- Add three to five small modular square fragments above the mark
- Avoid detailed hammer imagery or literal flames

### Logo Color Use

| Element | Color |
|---|---|
| Main anvil | `#E1E6EA` |
| Spreadsheet frame | `#AEB8C2` |
| Active spreadsheet cells | `#DE4141` |
| Modular particles | `#667482`, `#DE4141` |

### Wordmark

- “Sheet” in `#F5F7F9`
- “Forge” in `#DE4141`
- Medium or semibold weight
- Slightly tightened letter spacing
- No italics
- No decorative typography

---

# Typography

## Recommended Font Stack

### Primary Recommendation

**Inter**

Use Inter for all interface text because it is highly readable at small sizes and works well in structured productivity applications.

### Alternatives

- Geist
- Manrope
- IBM Plex Sans
- SF Pro Display and SF Pro Text for Apple-focused interfaces

## Type Scale

| Style | Size | Weight | Line Height | Use |
|---|---:|---:|---:|---|
| Display | 36 px | 700 | 42 px | Welcome screen app name |
| H1 | 28 px | 700 | 34 px | Product title |
| H2 | 22 px | 650 | 28 px | Screen title |
| H3 | 18 px | 650 | 24 px | Section heading |
| Body Large | 16 px | 500 | 24 px | Tagline and descriptive copy |
| Body | 14 px | 450 | 20 px | Standard interface text |
| Label | 13 px | 600 | 18 px | Buttons and navigation |
| Metadata | 12 px | 450 | 16 px | Updated dates and helper text |
| Micro | 11 px | 500 | 14 px | Small chips and captions |

### Typography Rules

- Use sentence case for headings and controls
- Avoid all caps
- Keep headings concise
- Limit body copy width to approximately 30 to 38 characters on mobile
- Use tabular numerals for prices, module counts, and progress indicators

---

# Spacing System

Use an 8-point spacing system with selective 4-point increments.

| Token | Size | Use |
|---|---:|---|
| Space 1 | 4 px | Icon-to-label micro spacing |
| Space 2 | 8 px | Tight component spacing |
| Space 3 | 12 px | Internal card spacing |
| Space 4 | 16 px | Standard component padding |
| Space 5 | 20 px | Section grouping |
| Space 6 | 24 px | Screen gutters |
| Space 8 | 32 px | Major section separation |
| Space 10 | 40 px | Hero spacing |
| Space 12 | 48 px | Large welcome-screen spacing |

### Mobile Screen Gutters

- Horizontal screen padding: `20–24 px`
- Top safe-area spacing: `16–24 px`
- Bottom navigation clearance: `80–96 px`

---

# Shape and Radius System

| Component | Radius |
|---|---:|
| Mobile device frame | 36–42 px |
| Primary button | 12 px |
| Cards | 10–14 px |
| Inputs | 10 px |
| Chips | 999 px |
| Icon containers | 10–12 px |
| Toggle track | 999 px |
| Bottom navigation | 0–16 px depending on implementation |

Use consistent medium radii. Avoid overly rounded, bubble-like components.

---

# Border and Divider System

## Borders

- Default border: `1 px solid #2D4053`
- Elevated border: `1 px solid #3A4D60`
- Active border: `1 px solid #DE4141`
- Disabled border: `1 px solid #293A4B`

## Dividers

- Divider color: `#314356`
- Opacity: 70–85%
- Use sparingly between dense sections
- Prefer spacing over dividers where possible

---

# Shadow and Elevation

Dark interfaces should use subtle shadows and edge contrast rather than heavy drop shadows.

## Card Shadow

```css
box-shadow:
  0 8px 24px rgba(0, 0, 0, 0.18),
  inset 0 1px 0 rgba(255, 255, 255, 0.025);
```

## Floating Panel Shadow

```css
box-shadow:
  0 14px 36px rgba(0, 0, 0, 0.28),
  0 2px 6px rgba(0, 0, 0, 0.18);
```

## Device Mockup Shadow

```css
box-shadow:
  0 24px 56px rgba(0, 0, 0, 0.38);
```

---

# Buttons

## Primary Button

- Background: `#DE4141`
- Hover: `#E64A45`
- Pressed: `#C93636`
- Text: `#FFFFFF`
- Height: `52–56 px`
- Radius: `12 px`
- Horizontal padding: `20–24 px`
- Font: 15–16 px, semibold
- Optional right-arrow icon

## Secondary Button

- Background: transparent or `#102338`
- Border: `1 px solid #536577`
- Text: `#F5F7F9`
- Hover background: `#172A3D`

## Text Button

- Text: `#ED5A53`
- No background
- Use for “View all,” edit actions, and low-priority controls

## Disabled Button

- Background: `#263747`
- Text: `#708091`
- No shadow

---

# Cards

## Standard Card

- Background: `#102338`
- Border: `#2D4053`
- Radius: `12 px`
- Padding: `14–16 px`

## Elevated Card

- Background: `#14283C`
- Border: `#3A4D60`
- Shadow: subtle elevation shadow
- Use for selected packs, active modules, or featured content

## Card Structure

1. Leading icon
2. Primary label
3. Supporting metadata
4. Status, action, or chevron
5. Optional progress indicator

---

# Iconography

## Style

- Rounded geometric line icons
- Stroke width: `1.75–2 px`
- Default size: `20–24 px`
- Small size: `16–18 px`
- Large product icon: `40–48 px`
- Avoid filled icons unless used inside colored product tiles

## Recommended Libraries

- Lucide
- Phosphor Icons
- Material Symbols Rounded

## Icon Colors

- Default: `#D8DEE4`
- Secondary: `#8794A1`
- Active: `#DE4141`
- Success: `#43B875`
- Disabled: `#667482`

---

# Navigation

## Top Navigation

- Height: `56–64 px`
- Left: menu or back arrow
- Center or left-aligned title
- Right: notification, help, share, bookmark, or profile control
- Background: `#0A1B2E`
- Bottom border: `#24384C`

## Bottom Navigation

- Height: `68–76 px`
- Background: `#0A1B2E`
- Top border: `#24384C`
- Five items maximum
- Active state uses red icon and label
- Inactive state uses grey

Recommended items:

- Home
- Packs
- Automations
- Workbooks
- Profile

---

# Screen Breakdown

## 1. Welcome Screen

### Purpose

Communicate the product’s value immediately and establish the brand.

### Structure

- Centered Sheet Forge logo
- Product name
- Thin red accent line
- Primary headline
- Supporting description
- Primary CTA
- Secondary CTA
- Minimal geometric grid pattern at bottom

### Recommended Copy

**Headline**

> Reusable spreadsheet automation you configure, not code

**Supporting text**

> Modular macro packs for any workbook.

### Background

- Main background: `#071525`
- Optional bottom pattern: `#24384C` at 30–45% opacity
- No photography
- No complex illustration beyond the logo system

---

## 2. Home Screen

### Primary Sections

#### Recent Workbooks

Each row should include:

- Spreadsheet icon
- Workbook name
- Last updated timestamp
- Status text and indicator

Examples:

- Sales Workbook
- Finance Q2 Close
- Inventory Tracker

#### Macro Packs

Use horizontally scrollable cards containing:

- Pack icon
- Pack name
- Module count
- Installed or available status

Examples:

- Finance Close Pack
- Inventory Audit
- Reconciliation Pack

#### Quick Actions

Four compact actions:

- Configure automation
- Add workbook
- Import pack
- Run automation

#### Templates

Use smaller horizontal cards:

- Monthly Finance Report
- Budget Planner
- Inventory Template

---

## 3. Macro Pack Product Page

### Example Product

**Finance Close Pack**

### Content Structure

- Back navigation
- Share and bookmark controls
- Pack icon
- Product title
- Status badge
- Short description
- Module count
- Estimated setup time
- Skill or complexity level
- Included feature list
- Compatibility section
- Price
- Add Pack CTA
- Guarantee or support note

### Example Feature List

- Trial balance validation
- Accruals and adjustments
- Reconciliation automation
- Financial report generation
- Close checklist and sign-off
- Audit trail and documentation

### Pricing Treatment

- Large price: `28–34 px`
- Small billing qualifier: `12–14 px`
- Price left-aligned
- CTA right-aligned or full-width below on smaller screens

---

## 4. Workbook Setup / Automation Builder

### Purpose

Make the no-code configuration workflow feel controlled and approachable.

### Step Structure

1. Select workbook
2. Choose modules
3. Configure modules
4. Review and apply

### Expanded Module Controls

Each module row contains:

- Module icon
- Module name
- Toggle
- Settings icon

Example modules:

- Trial Balance Validation
- Accruals and Adjustments
- Reconciliation Automation
- Financial Report Generation
- Audit Trail and Documentation

### Step Indicator

- Active step: red circle
- Completed step: muted red or check icon
- Future step: grey circle
- Connector: dotted or thin vertical line

### Primary Action

**Apply to Workbook**

- Sticky bottom button
- Full-width
- Red background
- Right-arrow icon

---

# Form Controls

## Toggle

- Width: `38–42 px`
- Height: `22–24 px`
- Active track: `#DE4141`
- Inactive track: `#536273`
- Thumb: `#F5F7F9`

## Input Field

- Height: `46–50 px`
- Background: `#102338`
- Border: `#33495D`
- Focus border: `#DE4141`
- Placeholder: `#667482`
- Text: `#F5F7F9`

## Dropdown

- Same visual styling as inputs
- Right-aligned chevron
- Open state uses elevated navy panel

## Checkbox

- Default border: `#667482`
- Selected background: `#DE4141`
- Checkmark: white

---

# Badges and Status Pills

## Best Seller

- Background: `#3B2028`
- Text: `#ED5A53`

## Installed

- Background: `#183B32`
- Text: `#43B875`

## Available

- Background: `#263747`
- Text: `#AEB8C2`

## Draft

- Background: transparent
- Text: `#8794A1`
- Optional grey status dot

Badge height should remain between `22–26 px`.

---

# Motion Guidelines

Use motion to communicate configuration and progress, not decoration.

| Interaction | Duration |
|---|---:|
| Button hover | 120 ms |
| Toggle change | 160 ms |
| Card expansion | 180–220 ms |
| Page transition | 220–280 ms |
| Modal opening | 180–240 ms |

Recommended easing:

```css
cubic-bezier(0.2, 0.8, 0.2, 1)
```

Avoid bounce effects.

---

# Accessibility

- Minimum body text contrast: WCAG AA
- Primary text on navy should use `#F5F7F9`
- Avoid grey text below `#8794A1` for essential information
- Minimum touch target: `44 × 44 px`
- Never communicate status by color alone
- Toggles should include labels and accessible state text
- Active red text on dark navy should generally use `#ED5A53`, not darker red values
- Support system font scaling where possible

---

# Design Tokens

```css
:root {
  --sf-bg-primary: #071525;
  --sf-bg-secondary: #0A1B2E;
  --sf-surface-primary: #102338;
  --sf-surface-secondary: #172A3D;
  --sf-surface-elevated: #14283C;

  --sf-border-default: #2D4053;
  --sf-border-strong: #3A4D60;
  --sf-divider: #314356;

  --sf-accent-primary: #DE4141;
  --sf-accent-hover: #E64A45;
  --sf-accent-pressed: #C93636;
  --sf-accent-soft: #3B2028;

  --sf-text-primary: #F5F7F9;
  --sf-text-secondary: #AEB8C2;
  --sf-text-muted: #8794A1;
  --sf-text-disabled: #667482;

  --sf-success: #43B875;
  --sf-success-soft: #183B32;
  --sf-blue: #566FAE;
  --sf-purple: #7855A8;
  --sf-warning: #D49A48;

  --sf-radius-sm: 8px;
  --sf-radius-md: 10px;
  --sf-radius-lg: 12px;
  --sf-radius-xl: 16px;
  --sf-radius-pill: 999px;

  --sf-space-1: 4px;
  --sf-space-2: 8px;
  --sf-space-3: 12px;
  --sf-space-4: 16px;
  --sf-space-5: 20px;
  --sf-space-6: 24px;
  --sf-space-8: 32px;
  --sf-space-10: 40px;
  --sf-space-12: 48px;
}
```

## Core Palette Summary

```text
Primary background:   #071525
Secondary background: #0A1B2E
Card surface:         #102338
Elevated surface:     #14283C
Primary red:          #DE4141
Red hover:            #E64A45
Primary text:         #F5F7F9
Secondary text:       #AEB8C2
Muted text:           #8794A1
Primary border:       #2D4053
Success:              #43B875
```

> These values are visually matched to the generated concept rather than extracted from a formal source design file, so slight differences may appear between the image and production implementation.
