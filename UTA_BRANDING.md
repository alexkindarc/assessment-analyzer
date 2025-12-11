# UTA Brand Implementation

The Assessment Report Analyzer has been styled according to UTA brand guidelines.

---

## Colors Used

### Primary Colors
| Color | Hex Code | Usage |
|-------|----------|-------|
| **UTA Blue** | `#0064b1` | Primary buttons, links, accents |
| **UTA Web Blue** | `#003865` | Sidebar background, headers, text |
| **UTA Orange** | `#F58025` | Accent bar, highlights, success states |
| **UTA Web Orange** | `#c45517` | Accessible text on light backgrounds |

### Implementation Notes
- Orange is used for **accents only** (not text), per UTA guidelines
- Sidebar uses dark blue (`#003865`) background with white text
- Primary action buttons use blue gradient
- Success messages have orange accent border
- Headers feature blue gradient with orange accent bar below

---

## Visual Elements

### Login Page
```
┌─────────────────────────────────────────────────────────────┐
│                                                             │
│         📊 Assessment Report Analyzer                       │
│    Office of Institutional Effectiveness and Reporting      │
│         The University of Texas at Arlington                │
│                                                             │
│         ┌───────────────────────────────────┐              │
│         │ ████████████████████████████████ │ ← Orange bar │
│         │                                   │              │
│         │  Enter password: [____________]   │              │
│         │                                   │              │
│         │  ℹ️ Contact administrator...       │              │
│         └───────────────────────────────────┘              │
│                                                             │
└─────────────────────────────────────────────────────────────┘
```

### Main Interface
```
┌─────────────────────────────────────────────────────────────┐
│ ████████████████████████████████████████████████████████████│ ← Blue header
│ 📊 Assessment Report Analyzer                               │
│ Office of Institutional Effectiveness and Reporting         │
│ ████████████████████████████████████████████████████████████│ ← Orange accent
├────────────┬────────────────────────────────────────────────┤
│            │                                                │
│  SIDEBAR   │  Main Content Area                             │
│  (Dark     │                                                │
│   Blue     │  ┌────────────────────────────────────────┐   │
│   #003865) │  │  [Analyze Report] [Batch Import] [Config]  │ ← Tabs
│            │  └────────────────────────────────────────┘   │
│  📊 Assessment │                                           │
│  Analyzer     │  Upload Report                              │
│               │                                             │
│  UTA IE       │  [Select report type ▼]                    │
│  (Orange text)│                                             │
│               │  ┌─────────────────────────┐              │
│  ───────────  │  │  Drop files here        │ ← Blue border │
│               │  │  or click to upload     │              │
│  👤 User Mode │  └─────────────────────────┘              │
│               │                                             │
│  ───────────  │  [🔍 Analyze Report] ← Blue button         │
│               │                                             │
│  API Key:     │                                             │
│  [••••••••]   │                                             │
│               │                                             │
├───────────────┴────────────────────────────────────────────┤
│ ████████████████████████████████████████████████████████████│ ← Footer
│ Assessment Report Analyzer                                  │ (Dark Blue)
│ Office of Institutional Effectiveness and Reporting         │
│ The University of Texas at Arlington                        │
└─────────────────────────────────────────────────────────────┘
```

---

## Brand Compliance

✅ **Primary colors dominant** - Blue and orange are the main colors
✅ **Orange for accents only** - Not used for text per guidelines
✅ **Accessible contrast** - Using `#003865` for dark text, white on dark backgrounds
✅ **Professional appearance** - Clean, modern design appropriate for university
✅ **Consistent with UTA web standards** - Following digital design system

---

## CSS Variables Reference

```css
:root {
    --uta-blue: #0064b1;
    --uta-web-blue: #003865;
    --uta-orange: #F58025;
    --uta-web-orange: #c45517;
    --uta-light-gray: #f5f7fa;
}
```

---

## Font

The app uses **sans-serif** system fonts, which provide:
- Clean, readable text
- Fast loading (no external font downloads)
- Similar appearance to UTA's Frutiger font family
