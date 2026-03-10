# Word-Autosave-Mac

**💾 Word for Mac Global AutoSave Plugin**
A lightweight VBA-powered add-in for Microsoft Word on macOS. It creates a global 10-minute autosave loop for all your documents, ensuring you never lose work on local or synced files (Dropbox, Google Drive, iCloud) without needing to save files as .docm.



**🚀 Features**
Global Scope: Works on all .docx and .doc files automatically.

Non-Invasive: Lives in your startup folder, not in your individual documents—perfect for sharing files with others without macro warnings.

Status Updates: Displays a brief "Auto-Save Performed" message in the Word status bar (bottom-left).

Smart Logic: Automatically ignores unsaved "Document1" drafts to prevent annoying "Save As" pop-ups.



**📥 Installation Methods**
Choose one of the following three methods to install the plugin.

Method 1: The "Set and Forget" Way (Recommended)
Moving the file to the Word Startup folder ensures it loads every time the application launches and won't be accidentally deleted during a desktop cleanup.

Open Finder.

Press Cmd + Shift + G.

Paste the following path:
~/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word

Drop AutoSavePlugin.dotm into this folder.

Restart Word.

Method 2: The "Fastest" Way (Terminal)
If you downloaded the file to your Downloads folder, run this command to move it to the Startup folder instantly:

Bash
mv ~/Downloads/AutoSavePlugin.dotm ~/Library/Group\ Containers/UBF8T346G9.Office/User\ Content.localized/Startup.localized/Word/

**Method 3: The "Easy" Way (Word Menu)**
Use this if you prefer using the Word interface. Note: Do not move the file after doing this, or the plugin will break.

Open Microsoft Word.

Go to Tools > Add-ins...

Click Add... and select your AutoSavePlugin.dotm file.

Ensure the box next to the filename is checked and click OK.



**🛠 Troubleshooting: "Grant Access" & Security**
1. The macOS Sandbox Prompt
Because macOS "sandboxes" applications for security, Word may occasionally ask for permission to save a file in a specific folder for the first time.

The Issue: You see a popup saying "Word needs additional permissions to save this file."

The Fix: Click Grant Access. Once you do this for a specific folder (like your "Documents" or "Work" folder), macOS will remember the permission and won't ask again for files in that location.

2. Enabling Macros
The first time you launch Word after installation, you may see a security warning.

The Fix: Since this is your own plugin (stored in your Library), click Enable Macros. To avoid this in the future, you can go to Word > Preferences > Security and adjust your macro settings, but "Enable" is the safest way to ensure the script runs.

3. "Document1" Not Saving
The script is intentionally designed not to autosave a file that has never been saved before. This prevents Word from constantly opening a "Save As" window while you are just jotting down quick notes.

The Fix: Save your document manually once to give it a name and a location. The 10-minute timer will take over from there.

**⚙️ Customization**
To change the save interval (e.g., to 5 minutes), open the VBA Editor (Opt + F11), locate the AutoSavePlugin project, and change the TimeValue in the code:

VBA
' Change "00:10:00" (10 mins) to "00:05:00" (5 mins)
Application.OnTime When:=Now + TimeValue("00:10:00"), Name:="GlobalAutoSave"



**✅ How to Verify it's Working**
The Welcome Message: The very first time you launch Word after installation, you will see a popup confirming the "AutoSave Plugin Successfully Installed."

The Status Bar: Every 10 minutes, look at the bottom-left corner of your Word window. You will see a brief message: 💾 Auto-Save Performed at [Time].

The Manual Test: * Open a saved document.

Press Opt + F8 (or go to Developer > Macros).

If you see GlobalAutoSave in the list, the plugin is active and running.
