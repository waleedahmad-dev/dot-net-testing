# Word Highlighter - Grammarly Clone

A beautiful Windows desktop application with a modern, Grammarly-inspired interface that integrates with Microsoft Word to highlight words and provide an interactive replacement widget. Click on any highlighted word and use the elegant popup widget to replace it instantly.

## ✨ Features

🎨 **Beautiful Modern UI**: Grammarly-inspired interface with card-based design, smooth shadows, and professional styling  
🔗 **Seamless Word Integration**: Connects effortlessly to Microsoft Word documents  
✨ **Smart Highlighting**: Automatically highlights all words with yellow background  
🖱️ **Click-to-Replace**: Click any highlighted word to open an elegant replacement widget  
⚡ **Quick Transformations**: Uppercase, lowercase, and capitalize with one click  
📊 **Activity Feed**: Real-time logging of all your replacements and actions  
🎯 **Live Detection**: Automatically detects when you click on highlighted words  
💅 **Smooth Animations**: Modern UI with rounded corners, card shadows, and hover effects

## Prerequisites

Before running this application, ensure you have:

1. **Windows Operating System** (Windows 10 or later recommended)
2. **Microsoft Office Word** installed (any recent version)
3. **.NET 8.0 SDK** or later
   - Download from: https://dotnet.microsoft.com/download/dotnet/8.0
4. **Visual Studio 2022** (optional, for development)
   - Community Edition is free: https://visualstudio.microsoft.com/

## Installation

### Option 1: Build from Source

1. **Clone or download this repository**

   ```powershell
   cd "c:\Users\Admin\Documents\Hobby\grammarly-clone"
   ```

2. **Restore NuGet packages**

   ```powershell
   dotnet restore
   ```

3. **Build the project**

   ```powershell
   dotnet build --configuration Release
   ```

4. **Run the application**
   ```powershell
   dotnet run
   ```

### Option 2: Run in Debug Mode

```powershell
dotnet run --project WordHighlighter.csproj
```

## Usage

### Step-by-Step Guide

1. **Open Microsoft Word**

   - Launch Microsoft Word
   - Open an existing document or create a new one with some text

2. **Start the Word Highlighter Application**

   - Run the application using one of the installation methods above
   - You'll see a beautiful, modern interface with cards and professional styling

3. **Connect to Word**

   - Click the green "Connect to Word" button in the welcome card
   - The connection status badge will turn green and say "Connected"
   - The highlighting action cards will become enabled

4. **Highlight All Words**

   - Click "Highlight All" in the blue action card
   - All words in your Word document will be highlighted in yellow
   - A confirmation shows the number of words highlighted

5. **Replace a Word**

   - In Microsoft Word, click on any highlighted word
   - A beautiful replacement widget appears with smooth shadow effects
   - View the original word and enter your replacement
   - Use quick transformation buttons (🔤 UPPERCASE, 🔡 lowercase, 🔠 Capitalize)
   - Click "Replace Word" or press Enter

6. **Clear Highlights**
   - Click "Clear All" in the red action card to remove all highlighting
   - Your document text and replacements remain unchanged

## Features in Detail

### Main Window

- **Modern Header**: Grammarly-inspired design with logo badge and app title
- **Connection Status Badge**: Color-coded indicator (yellow when disconnected, green when connected)
- **Welcome Card**: Large, prominent card with connection instructions and icon
- **Action Cards**: Beautiful cards for Highlight and Clear operations with icons
- **How It Works**: Step-by-step guide with numbered badges
- **Activity Feed**: Real-time activity log with emoji indicators
- **Card-based Design**: Clean white cards with subtle shadows on light gray background

### Replacement Widget

- **Borderless Modern Design**: Smooth rounded corners with elegant drop shadow
- **Original Word Display**: Clear display of the word you clicked
- **Replacement Input**: Large, focused text input with border highlighting
- **Quick Transformations**:
  - **🔤 UPPERCASE**: Convert to UPPERCASE
  - **🔡 lowercase**: Convert to lowercase
  - **🔠 Capitalize**: Convert To Title Case
- **Action Buttons**: Modern buttons with hover effects
- **Error Display**: Friendly error messages in red alert boxes
- **Keyboard Shortcuts**:
  - `Enter`: Apply replacement
  - `Escape`: Cancel

### UI Design Elements

- **Color Palette**:
  - Primary Green: #15C39A (Grammarly-inspired)
  - Text Dark: #111827
  - Text Light: #6B7280
  - Background: #F5F7FA
- **Typography**: Modern, clean fonts with proper hierarchy
- **Shadows**: Subtle drop shadows for depth (0.08 opacity)
- **Rounded Corners**: 8-16px border radius for modern feel
- **Hover Effects**: Interactive feedback on all clickable elements
- **Emojis**: Visual indicators in activity log for better readability

## Technical Details

### Architecture

- **Framework**: .NET 8.0 WPF (Windows Presentation Foundation)
- **Office Integration**: Microsoft Office Interop Word
- **Language**: C#
- **UI Framework**: XAML

### Project Structure

```
grammarly-clone/
├── App.xaml                      # Application definition
├── App.xaml.cs                   # Application code-behind
├── MainWindow.xaml               # Main window UI
├── MainWindow.xaml.cs            # Main window logic
├── WordReplacementWidget.xaml    # Replacement popup UI
├── WordReplacementWidget.xaml.cs # Replacement popup logic
├── WordInteropService.cs         # Word COM interop service
├── WordHighlighter.csproj        # Project file
└── app.manifest                  # Application manifest
```

### Key Components

1. **WordInteropService.cs**

   - Manages COM interop with Microsoft Word
   - Handles highlighting, selection detection, and word replacement
   - Implements IDisposable for proper COM object cleanup

2. **MainWindow**

   - Main application interface
   - Manages connection to Word
   - Monitors for word clicks using a timer
   - Displays activity log

3. **WordReplacementWidget**
   - Popup dialog for word replacement
   - Provides text transformation utilities
   - Keyboard-friendly interface

## Troubleshooting

### "Could not connect to Word"

- Ensure Microsoft Word is running
- Make sure you have a document open in Word
- Try restarting Word and the application

### "Access Denied" or COM Errors

- Run the application with appropriate permissions
- Ensure Word is not in protected mode
- Check that COM automation is enabled in Word

### Highlights not appearing

- The document might be in read-only mode
- Try saving the document first
- Check Word's macro security settings

### Click detection not working

- Ensure you're clicking directly on the highlighted text
- The word must have yellow highlighting
- Try selecting the word by clicking and holding

## Building for Distribution

To create a standalone executable:

```powershell
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true
```

The executable will be in: `bin\Release\net8.0-windows\win-x64\publish\`

## Limitations

- Only works on Windows with Microsoft Word installed
- Requires Word to be open with a document
- Detects clicks through selection monitoring (500ms polling interval)
- COM interop may have performance considerations with very large documents

## Future Enhancements

- [ ] Grammar and spell-checking suggestions with AI
- [ ] Thesaurus integration for synonym suggestions
- [ ] Multiple highlighting colors for different word categories
- [ ] Support for specific word categories (nouns, verbs, adjectives)
- [ ] Dictionary definitions popup in the widget
- [ ] Undo/redo functionality with history
- [ ] Support for multiple simultaneous Word documents
- [ ] Customizable keyboard shortcuts
- [ ] Dark mode theme
- [ ] Animated transitions and micro-interactions
- [ ] Word statistics dashboard
- [ ] Custom suggestion algorithms

## License

This project is provided as-is for educational and personal use.

## Support

For issues or questions:

1. Check the Troubleshooting section above
2. Ensure all prerequisites are installed
3. Verify Microsoft Word is properly installed and functioning

## Contributing

This is a demonstration project. Feel free to fork and modify for your needs.

---

**Note**: This application requires Microsoft Word to be installed and will interact with open Word documents. Always save your work before using automated text manipulation tools.
