#pragma once

#include <Guiddef.h>

#ifdef INITGUID
# error INITGUID should not be declared when using DeclareGuid.h.
#endif

#define DECLARE_GUID(name) DEFINE_GUID(name, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
