# PowerPoint Add-in for Smart PowerPoint

## Overview

This directory contains the Office Add-in that allows professors to use Smart PowerPoint directly within Microsoft PowerPoint.

## Structure

```
powerpoint-addin/
├── manifest.xml          # Add-in configuration
├── taskpane.html        # Main UI (embedded in PowerPoint)
├── taskpane.js          # Connection logic
├── taskpane.css         # Styling
└── README.md            # This file
```

## Features

- **Auto Slide Detection** - Detects current slide automatically
- **Real-time Notifications** - Shows student doubts in side panel
- **WebSocket Connection** - Connects to existing Smart PowerPoint backend
- **Seamless Integration** - Works within PowerPoint presentation mode

## Installation (Development)

### Prerequisites

- Microsoft PowerPoint (Windows or Mac)
- Node.js installed (for local server)
- Smart PowerPoint backend running

### Steps

1. **Start the add-in server:**

   ```bash
   cd powerpoint-addin
   python3 -m http.server 3000
   ```

2. **Sideload the add-in:**

   **Windows:**

   - File → Options → Trust Center → Trust Center Settings
   - Trusted Add-in Catalogs → Add `C:\path\to\powerpoint-addin` as trusted location
   - Copy `manifest.xml` to the trusted catalog folder
   - Restart PowerPoint
   - Insert → My Add-ins → Smart PowerPoint

   **Mac:**

   - Create folder: `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`
   - Copy `manifest.xml` to this folder
   - Restart PowerPoint
   - Insert → My Add-ins → Smart PowerPoint

3. **Start presenting:**
   - Open your PowerPoint presentation
   - Click "Smart PowerPoint" in the ribbon
   - Task pane will open on the right
   - Start presenting (F5) - the add-in will track your slides!

## How It Works

1. **Slide Change Detection:**

   ```javascript
   Office.context.document.addHandlerAsync(
     Office.EventType.DocumentSelectionChanged,
     onSlideChange
   );
   ```

2. **WebSocket Connection:**
   Connects to `ws://localhost:8000/ws/{session_id}` (your existing backend)

3. **Notification Display:**
   Shows the same glass-morphism notifications from the web app

## Production Deployment

To publish to Microsoft AppSource:

1. Update `manifest.xml` with production URLs
2. Host add-in files on HTTPS server
3. Submit to Microsoft for review
4. Users can install from Office Store

## Development Tips

- Use PowerPoint Online for faster testing (no sideloading needed)
- Debug with F12 Developer Tools (right-click in task pane)
- Test on both Windows and Mac PowerPoint
- Ensure CORS is enabled on backend for localhost:3000

## Next Steps

1. **Test with real presentation** - Use your actual slides
2. **Add slide content extraction** - Send slide text to backend
3. **Improve error handling** - Better connection status
4. **Add settings panel** - Let professor configure session ID
