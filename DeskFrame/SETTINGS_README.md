# Global Settings System

## Overview
Global settings provide **default values for newly created frames**. These settings do NOT affect existing frames.

## Files
- `AppSettings.cs` - Settings model and JSON management
- `SettingsWindow.xaml` / `.xaml.cs` - Global settings UI
- `appsettings.json` - Stored in `%AppData%/DeskFrame/`

## Settings

### Opacity (0-100%)
- Controls the transparency of the frame background
- Default: 10% (26/255 in byte value)
- Applied via Alpha channel in ARGB color

### Background Color
- Hex color code (e.g., `#0C000000`)
- Default: `#0C000000` (very dark with low opacity)
- Preset colors available in UI

## How It Works

### 1. Global Settings (Tray Menu ? Settings)
- Opens `SettingsWindow`
- Displays current global defaults
- Changes saved to `appsettings.json`
- **Only affects NEW frames created after saving**
- **Does NOT modify existing frames**

### 2. Frame-Specific Settings (?? button on each frame)
- Opens `FrameSpecificSettingsWindow`
- Modifies only that specific frame
- Saved to Windows Registry per frame
- Independent of global settings

### 3. New Frame Creation
When a new frame is created:
1. First loads values from Registry (if they exist)
2. Falls back to `AppSettings` for missing values
3. Creates instance with combined settings

## Usage

### Setting Global Defaults
```
1. Right-click Tray Icon ? Settings
2. Adjust Opacity and Background Color
3. Click "Save"
4. All NEW frames will use these defaults
```

### Customizing Individual Frames
```
1. Click ?? on any frame's title bar
2. Adjust settings for that frame
3. Click "Apply"
4. Only that frame is affected
```

## Technical Details

### Priority Order (New Frames)
1. Registry values (if exist from previous sessions)
2. AppSettings defaults (if no Registry value)
3. Hard-coded defaults (as fallback)

### Storage Locations
- **Global**: `%AppData%/DeskFrame/appsettings.json`
- **Per Frame**: `HKEY_CURRENT_USER\SOFTWARE\DeskFrame\Instances\{FrameName}`

## Code Flow

### Global Settings Save
```csharp
SettingsWindow.SaveButton_Click()
  ?
AppSettings.Save()  // Writes to JSON
  ?
// Does NOT update existing frames
```

### Frame Creation
```csharp
Instance Constructor
  ?
Load from Registry (if exists)
  ?
Load from AppSettings (if Registry empty)
  ?
Apply to new frame
```

### Frame-Specific Update
```csharp
FrameSpecificSettingsWindow.ApplyButton_Click()
  ?
Update Instance properties
  ?
Call frame.ChangeBackgroundOpacity()
  ?
Save to Registry
```

## Important Notes

?? **Global Settings DO NOT Affect Existing Frames**
- Changing global settings only affects new frames
- To update an existing frame, use its ?? button

? **Frame-Specific Settings Are Independent**
- Each frame can have unique settings
- Saved separately in Registry
- Not affected by global setting changes

?? **Best Practice**
1. Set global defaults first (for new frames)
2. Create frames (they inherit global defaults)
3. Customize individual frames as needed (via ??)

## Migration from Old System
If you have frames created before this update:
- They retain their existing settings
- Use ?? button to update opacity/color
- New frames will use global defaults
