CustomShortcutFileType
This is sample implementation of custom windows shortcut .customlnk similar to (.lnk or .appref-ms)

What is working
The shortcut is working in explorer, icon is shown and double click executes the command
The shortcut is shown in Windows10 start menu (when added to %USERPROFILE%\AppData\Roaming\Microsoft\Windows\Start Menu\Programs)
The icon is also shown in W10 start menu - for this to work, registration is setting CLSID class guid owner to NT Service\TrustedInstaller

How to use
0.Open a Visual Studio Command Prompt and go to the project directory, create CustomShortcutLibrary.tlb with this command:
  midl CustomShortcutLibrary.idl
1.Build either the x86 or x64 platform based on the platform of your Windows installation.
2.Run command prompt as Administrator and register the resulting build output: regsvr32.exe "C:\Path\To\CustomShortcutFileType.dll"
3.Create a test shortcut of the new .customlnk file type:
   -Select an EXE to be launched when the shortcut is activated. This can be anything on your system, such as C:\WINDOWS\System32\calc.exe
   -Select an ICO to be displayed when Windows Explorer is showing the shortcut. On my system, I found that Visual Studio had installed a plaethora of icons at C:\Program Files (x86)\Microsoft Visual Studio\common\Graphics\Icons. Your mileage may vary, but freely downloadable ICO files can be easily found with simple Google searches if necessary.
   -Open Notepad, and place the full path to the EXE on the first line and the full path to the ICO on the second line.
   -Save the file, specifying the name (including the double-quotes) as e.g. "My Test Shortcut.customlnk", and selecting "UTF-8" for the encoding.
4. Restart explorer.exe

For details see:
https://stackoverflow.com/questions/39002571/windows-explorer-custom-shortcut-file-types
https://social.msdn.microsoft.com/Forums/windowsdesktop/en-US/7c88d220-2363-4cfb-8ae1-87be143d85b7/custom-shortcut-filetypes?forum=windowsgeneraldevelopmentissues

Notes:
For W10 Start menu icon IExtractIconW.GetIconLocation is called by process RuntimeBroker.exe but only if HKEY_CLASSES_ROOT\CLSID\{42465C3A-83D3-4310-B27D-F271DE372764} owner is "NT Service\TrustedInstaller"