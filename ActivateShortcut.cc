#include "CustomShortcut.h"
#include "Utilities.h"

#include <string>
#include <vector>

using namespace std;

void CALLBACK ActivateShortcut(HWND hWnd, HINSTANCE hinst, LPSTR lpszCmdLine, int nCmdShow)
{
  // Convert the command-line to a native string.
  wstring cmdLine;

  cmdLine.resize(MultiByteToWideChar(CP_ACP, 0, lpszCmdLine, -1, NULL, 0));
  cmdLine.resize(MultiByteToWideChar(CP_ACP, 0, lpszCmdLine, -1, &cmdLine[0], (int)cmdLine.size()));

  if ((cmdLine.size() > 0) && (cmdLine[cmdLine.size() - 1] == '\0'))
    cmdLine.erase(cmdLine.size() - 1, 1);

  // Load in the shortcut.
  CustomShortcut shortcut;

  if (!SUCCEEDED(shortcut.Load(cmdLine.c_str(), IGNORE)))
    return;

  // Extract the path to the EXE it points at.
  wstring targetPath;

  targetPath.resize(32768);

  if (!SUCCEEDED(shortcut.GetPath(&targetPath[0], targetPath.size(), NULL, SLGP_RAWPATH)))
    return;

  targetPath.resize(wcsnlen_s(&targetPath[0], targetPath.size()));

  // Launch the file (if found).
  if (FileExists(targetPath))
  {
    wstring commandLine = wstring(L"\"") + targetPath + L"\"";

    STARTUPINFOW startInfo;
    PROCESS_INFORMATION processInfo;

    memset(&startInfo, 0, sizeof(startInfo));

    startInfo.cb = sizeof(startInfo);

    BOOL success = CreateProcessW(
      targetPath.c_str(),
      &commandLine[0],
      NULL, // process security attributes
      NULL, // thread security attributes
      FALSE, // inherit handles (don't)
      CREATE_NEW_PROCESS_GROUP,
      NULL, // environment block (just use that of the parent process)
      NULL, // initial working directory (just use our current working directory)
      &startInfo,
      &processInfo);

    if (success)
    {
      CloseHandle(processInfo.hProcess);
      CloseHandle(processInfo.hThread);
    }
  }
}
