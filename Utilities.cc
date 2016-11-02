#include "Utilities.h"

#include <Windows.h>

#include <string>
#include <vector>

using namespace std;

wstring GetModulePath(void *hInstance/* = NULL*/)
{
  vector<wchar_t> fileNameBuffer;

  fileNameBuffer.resize(150);

  while (true)
  {
    DWORD charactersCopied = GetModuleFileNameW((HMODULE)hInstance, &fileNameBuffer[0], (DWORD)fileNameBuffer.size());

    if (charactersCopied < fileNameBuffer.size())
    {
      fileNameBuffer.resize(charactersCopied);
      break;
    }
    else
      fileNameBuffer.resize(fileNameBuffer.size() * 2);
  }

  return &fileNameBuffer[0];
}

bool FileExists(const wstring &filePath)
{
  DWORD attributes = GetFileAttributesW(filePath.c_str());

  return (attributes != INVALID_FILE_ATTRIBUTES) && ((attributes & FILE_ATTRIBUTE_DIRECTORY) == 0);
}
