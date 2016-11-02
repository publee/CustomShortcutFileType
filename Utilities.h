#pragma once

#include <string>

std::wstring GetModulePath(void *hInstance/* = NULL*/);
bool FileExists(const std::wstring &filePath);
