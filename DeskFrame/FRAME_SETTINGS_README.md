# Frame-Specific Settings Feature

## Overview
Each frame now has its own settings button in the title bar, allowing you to customize individual frames independently.

## Location
The **Settings** button (??) is located in the title bar, just left of the minimize/expand chevron button.

## Features

### Frame-Specific Settings Window
- **Opacity Control**: Adjust transparency for this frame only (0-100%)
- **Background Color**: Set a custom background color for this frame
  - Manual HEX input (#AARRGGBB format)
  - Color picker dialog
  - Quick preset colors (Black, Dark Gray, Blue, Green)
- **Live Preview**: See changes before applying
- **Apply Button**: Apply settings to this frame only
- **Reset to Global**: Restore this frame to use global settings

## How to Use

1. **Click the Settings button** (??) on any frame's title bar
2. **Adjust settings** using sliders and color inputs
3. **Preview** your changes in real-time
4. **Click "Apply"** to save settings to this frame
5. **Click "Reset to Global"** to revert to global settings

## Key Differences from Global Settings

| Feature | Global Settings | Frame-Specific Settings |
|---------|----------------|-------------------------|
| **Access** | Tray menu ? Settings | Frame title bar ? ?? button |
| **Scope** | All frames | Single frame only |
| **Button Text** | "Save" | "Apply" |
| **Reset** | "Reset to Default" | "Reset to Global" |
| **Persistence** | Saved to `app_settings.json` | Saved to Windows Registry per frame |

## Use Cases

### Example 1: Different Projects
- **Frame 1** (Work documents): Dark blue background, 90% opacity
- **Frame 2** (Personal files): Green background, 70% opacity
- **Frame 3** (Downloads): Black background, 100% opacity

### Example 2: Visual Organization
- Important folders: High opacity, bright colors
- Archive folders: Low opacity, muted colors
- Temporary folders: Minimal opacity, neutral colors

## Technical Details

- Settings are saved immediately to the frame's Instance object
- Changes persist across application restarts
- Each frame's settings are stored in Windows Registry under `HKCU\SOFTWARE\DeskFrame\Instances\[FrameName]`
- Settings include:
  - `Opacity`: 0-255 value
  - `ListViewBackgroundColor`: HEX color code

## Tips

- Use **global settings** to set defaults for new frames
- Use **frame-specific settings** to customize existing frames
- The preview panel shows exactly how your frame will look
- Settings are applied instantly when you click "Apply"
