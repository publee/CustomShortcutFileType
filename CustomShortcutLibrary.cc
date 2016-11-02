#include <Windows.h>

#include <InitGuid.h>

#include <atlbase.h>
#include <statreg.h>

#include "ClassFactory.h"
#include "Utilities.h"
#include "Resource.h"

#include <string>

#define SHORTCUT_FILE_EXTENSION L".customlnk"

using namespace std;

// {EE9A48C8-1B1F-4985-8DA8-459209092971}
DEFINE_GUID(LIBID_CustomShortcutLibrary,
  0xee9a48c8, 0x1b1f, 0x4985, 0x8d, 0xa8, 0x45, 0x92, 0x9, 0x9, 0x29, 0x71);

HINSTANCE g_hInstance;
LONG g_moduleCount;

extern "C" INT STDMETHODCALLTYPE DllMain(HINSTANCE hInstance, DWORD reason, void *)
{
  if (reason == DLL_PROCESS_ATTACH)
  {
    g_hInstance = hInstance;
    DisableThreadLibraryCalls(hInstance);
  }

  return 1;
}

extern "C" HRESULT STDMETHODCALLTYPE DllCanUnloadNow()
{
  return (g_moduleCount == 0) ? S_OK : S_FALSE;
}

extern "C" HRESULT STDMETHODCALLTYPE DllGetClassObject(REFCLSID rclsid, REFIID riid, void **ppv)
{
  if (ppv == NULL)
    return E_POINTER;

  HRESULT result = CLASS_E_CLASSNOTAVAILABLE;

  ClassFactory *factory = new ClassFactory();

  if (factory->SetClass(rclsid))
    result = factory->QueryInterface(riid, ppv);

  if (FAILED(result))
  {
    ppv = NULL;
    delete factory;
  }

  return result;
}

extern "C" HRESULT STDMETHODCALLTYPE DllRegisterServer()
{
  wstring filePath = GetModulePath(g_hInstance);

  const wchar_t *filePathStr = filePath.c_str();

  CRegObject regObject;

  regObject.FinalConstruct();

  regObject.AddReplacement(L"Module", filePathStr);
  regObject.AddReplacement(L"ShortcutFileExtension", SHORTCUT_FILE_EXTENSION);
  regObject.ResourceRegister(filePathStr, IDR_REGISTRY, L"REGISTRY");

  CComPtr<ITypeLib> pTypeLib;

  LoadTypeLib(filePathStr, &pTypeLib);

  HRESULT result = RegisterTypeLib(pTypeLib, filePathStr, NULL);

  if (FAILED(result))
    return SELFREG_E_TYPELIB;
  else
    return S_OK;
}

extern "C" HRESULT STDMETHODCALLTYPE DllUnregisterServer()
{
  wstring filePath = GetModulePath(g_hInstance);

  const wchar_t *filePathStr = filePath.c_str();

  CRegObject regObject;

  regObject.FinalConstruct();

  regObject.AddReplacement(L"Module", filePathStr);
  regObject.AddReplacement(L"ShortcutFileExtension", SHORTCUT_FILE_EXTENSION);
  regObject.ResourceUnregister(filePathStr, IDR_REGISTRY, L"REGISTRY");

  HRESULT result = UnRegisterTypeLib(LIBID_CustomShortcutLibrary, 1, 0, 0, SYS_WIN32);

  if (FAILED(result))
    return SELFREG_E_TYPELIB;
  else
    return S_OK;
}
