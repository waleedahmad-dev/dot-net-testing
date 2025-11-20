# Quick Start Guide

Get up and running with Word Highlighter in under 5 minutes!

## Prerequisites Check ✅

Before you begin, make sure you have:

- [ ] Windows 10 or later
- [ ] Microsoft Office Word installed
- [ ] .NET 8.0 SDK installed ([Download here](https://dotnet.microsoft.com/download/dotnet/8.0))

## Installation (2 minutes)

### Option 1: Quick Run (Recommended for first time)

1. Open PowerShell in the project folder:

   ```powershell
   cd "c:\Users\Admin\Documents\Hobby\grammarly-clone"
   ```

2. Restore dependencies:

   ```powershell
   dotnet restore
   ```

3. Run the application:
   ```powershell
   dotnet run
   ```

### Option 2: Build and Run

1. Build the project:

   ```powershell
   dotnet build --configuration Release
   ```

2. Run the executable:
   ```powershell
   .\bin\Release\net8.0-windows\WordHighlighter.exe
   ```

## First Use (3 minutes)

### Step 1: Prepare a Word Document

1. Open Microsoft Word
2. Create a new document or open an existing one
3. Add some text (at least a few sentences)

### Step 2: Connect the App

1. Launch Word Highlighter
2. You'll see a beautiful interface with a welcome card
3. Click the green **"Connect to Word"** button
4. The status badge should turn green ✅

### Step 3: Highlight Words

1. Click **"Highlight All"** in the blue card
2. Watch as all words in your document turn yellow
3. You'll see a confirmation with the word count

### Step 4: Replace Your First Word

1. Go back to Microsoft Word
2. Click on any highlighted word
3. A beautiful widget appears!
4. Type a replacement word
5. Try a quick transformation (🔤 UPPERCASE, 🔡 lowercase, 🔠 Capitalize)
6. Click **"Replace Word"** or press Enter

### Step 5: Explore

- Check the Activity Feed to see your changes logged
- Try replacing more words
- Use the **"Clear All"** button when done

## Tips & Tricks 💡

### Keyboard Shortcuts

- **Enter** - Apply replacement
- **Escape** - Cancel widget

### Quick Transformations

Use the transformation buttons to quickly change case without retyping

### Activity Feed

Watch the Activity Feed (bottom of main window) for real-time updates

### Multiple Replacements

You can replace as many words as you want - the widget appears for each click

## Troubleshooting

### "Could not connect to Word"

- **Solution**: Make sure Word is running and has a document open
- Try closing and reopening Word
- Restart the application

### "No words highlighted"

- **Solution**: Make sure you clicked "Highlight All" after connecting
- Check that your document has text content

### Widget doesn't appear when clicking

- **Solution**: Make sure you're clicking on a highlighted (yellow) word
- The word must be selected in Word
- Wait a moment, the detection runs every 500ms

### Application won't start

- **Solution**: Check that .NET 8.0 SDK is installed
- Run `dotnet --version` in PowerShell to verify
- Reinstall .NET if needed

## What's Next?

Now that you're familiar with the basics:

1. **Try different documents** - Test with longer documents
2. **Experiment with transformations** - Use the quick action buttons
3. **Monitor the activity log** - See all your changes in real-time
4. **Read the full README** - Learn about advanced features

## Need Help?

- Check the main [README.md](README.md) for detailed documentation
- Review [DESIGN.md](DESIGN.md) for UI/UX guidelines
- Look at the troubleshooting section in README

## Video Walkthrough (Coming Soon)

We're working on a video tutorial to show you everything in action!

---

**Enjoy using Word Highlighter! 🎉**

Made with ❤️ for productivity enthusiasts
