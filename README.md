# Thunderbird “Send to” Shortcut (Windows)

This repository contains a lightweight setup for adding **Thunderbird** to the Windows **Send to** menu so that you can select one or more files in File Explorer, 
right-click, and choose **Send to → Thunderbird**. 

A single Thunderbird compose window opens with all selected files attached, and no console window flashes.

## Instructions

1. Either copy and paste the below vbs code into a new text file and name it Thunderbird.vbs or download [ThunderBird.vbs](https://github.com/DisasterofPuppets/Add_Thunderbird_SendTo_Context_menu/blob/main/ThunderBird.vbs) from the code repo.

2. Save `Thunderbird.vbs` into into a stable location. (I put mine on my backup drive)


## VBScript contents

```vbscript
Set objShell = CreateObject("WScript.Shell")
Set objArgs  = WScript.Arguments

If objArgs.Count = 0 Then WScript.Quit

tbird = "C:\Program Files\Mozilla Thunderbird\thunderbird.exe"

' Build a comma-separated attachment list: attachment='file:///path1,file:///path2'
attachments = ""
For i = 0 To objArgs.Count - 1
    filePath = objArgs(i)
    fileUrl = "file:///" & Replace(filePath, "\", "/")
    fileUrl = Replace(fileUrl, " ", "%20")
    attachments = attachments & fileUrl
    If i < objArgs.Count - 1 Then attachments = attachments & ","
Next

command = """" & tbird & """ -compose " & """attachment='" & attachments & "'"""
objShell.Run command, 0, False   ' 0 = hidden window
```

3. Press `Win + R`, type `shell:sendto`, and press Enter. This opens your SendTo folder:
4. 
   ```
   C:\Users\<YOU>\AppData\Roaming\Microsoft\Windows\SendTo
   ```
   
5. Create a new shortcut named **Thunderbird** with these properties: (or grab [ThunderBird.lnk](https://github.com/DisasterofPuppets/Add_Thunderbird_SendTo_Context_menu/blob/main/ThunderBird.lnk) from the code files, paste, right click > properties, and update the link to your ThunderBird.vbs file
   
   - **Target:**  
     ```
     C:\Windows\System32\wscript.exe "Y:\The Folder\Where You Saved\ThunderBird.vbs"
     ```
   - **Start in:**  
     ```
     "C:\Program Files\Mozilla Thunderbird"
     ```
   - (Optional) Set the icon to the Thunderbird executable for clarity.

6. That’s it. Right-click any file(s) → **Send to → Thunderbird**.



### How it works
- **Single compose window:** Uses one `-compose` invocation with a comma-separated attachment list, which Thunderbird opens as a single draft.
- **No console flash:** `wscript.exe` plus the `0` window style hides helper windows.
- **Spaces handled:** Percent-encoding keeps attachment URLs valid.

## Notes and variations
- You can store the VBScript on another drive; just update the shortcut target to match.
- If you need this for all users, place the shortcut in each profile’s SendTo folder or deploy via a script that writes the `.lnk` file.
- A registry key cannot populate the SendTo menu by itself; the SendTo menu is driven by the contents of the SendTo folder. A context-menu verb (registry) is a different mechanism and appears higher in the right-click menu.
