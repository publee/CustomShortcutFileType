#pragma once

#include <Windows.h>

// Without <InitGuid.h>, these definitions are actually declarations.
#include "DeclareGuid.h"

DECLARE_GUID(LIBID_CustomShortcutLibrary);
DECLARE_GUID(CLSID_CustomShortcut);

extern HINSTANCE g_hInstance;
extern LONG g_moduleCount;
