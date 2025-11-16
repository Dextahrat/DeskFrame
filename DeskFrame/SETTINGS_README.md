# Settings System

## Overview
DeskFrame now includes a modernized settings system that allows you to customize the appearance of your frames.

## Settings Location
Settings are stored in: `%AppData%\DeskFrame\app_settings.json`

## Available Settings

### 1. Default Opacity
- **Range**: 0% - 100%
- **Description**: Controls the transparency level of frames
- **Default**: 85%
- **Note**: When changed, applies to ALL existing and new frames immediately

### 2. Background Color
- **Format**: `#AARRGGBB` (Hexadecimal with Alpha channel)
- **Description**: Sets the background color for frames
- **Default**: `#0C000000` (semi-transparent black)
- **Note**: When changed, applies to ALL existing and new frames immediately
- **Common Presets**:
  - Black: `#0C000000`
  - Dark Gray: `#0C202020`
  - Blue: `#0C001F3F`
  - Green: `#0C0F2F1F`

## How to Use

1. **Open Settings**: Click the "Settings" button in the tray menu
2. **Adjust Opacity**: Use the slider to set transparency (0-100%)
3. **Choose Background Color**: 
   - Enter a hex color code directly
   - Use the "Pick" button to open a color picker
   - Click preset color buttons for quick selection
4. **Preview**: See your changes in real-time in the preview panel
5. **Save**: Click "Save" to apply settings to **ALL frames** (existing + new)
6. **Reset**: Click "Reset to Default" to restore default values

## Important Notes

- ? Settings apply to **ALL FRAMES** immediately upon saving
- ? Both existing open frames and new frames will use the new settings
- ? Changes are saved to file and persisted across application restarts
- ? The preview panel shows exactly how your frames will look
- ? Settings are automatically loaded when the application starts

## Technical Details

When you click "Save":
1. Settings are saved to `app_settings.json`
2. All currently open frames are updated immediately
3. The `Instance.Opacity` and `Instance.ListViewBackgroundColor` properties are updated
4. The UI is refreshed to show the new appearance
5. New frames created after saving will also use these settings
