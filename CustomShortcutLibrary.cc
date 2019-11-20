#include <Windows.h>
#include <aclapi.h>

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

DWORD SecuritySetTrustedInstallerOwner(HKEY hKey, LPCTSTR lpSubKey);
void EnableTakeOwnershipPrivilege();

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
  if (FAILED(result)) return SELFREG_E_TYPELIB;

  EnableTakeOwnershipPrivilege();
  DWORD dwRes = SecuritySetTrustedInstallerOwner(HKEY_CLASSES_ROOT, L"CLSID\\{42465C3A-83D3-4310-B27D-F271DE372764}");
  if (dwRes != ERROR_SUCCESS) return SELFREG_E_TYPELIB;

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


DWORD SecuritySetTrustedInstallerOwner(HKEY hKey, LPCTSTR lpSubKey)
{
  DWORD ret;
  HKEY hk;
  DWORD dwSize = 0;
  PSECURITY_DESCRIPTOR pOldSD = NULL;
  PSECURITY_DESCRIPTOR pNewSD = NULL;
  PSID psid = NULL;
  do
  {
    ret = RegOpenKeyEx(hKey, lpSubKey, 0, KEY_READ | READ_CONTROL | WRITE_OWNER, &hk);
    if (ret != ERROR_SUCCESS)
    {
      break;
    }

    RegGetKeySecurity(hk, DACL_SECURITY_INFORMATION | OWNER_SECURITY_INFORMATION, NULL, &dwSize);
    pOldSD = (PSECURITY_DESCRIPTOR)LocalAlloc(LPTR, dwSize);
    ret = RegGetKeySecurity(hk, DACL_SECURITY_INFORMATION | OWNER_SECURITY_INFORMATION, pOldSD, &dwSize);
    if (ret != ERROR_SUCCESS)
    {
      break;
    }

    pNewSD = (PSECURITY_DESCRIPTOR)LocalAlloc(LPTR, SECURITY_DESCRIPTOR_MIN_LENGTH);
    if (!InitializeSecurityDescriptor(pNewSD, SECURITY_DESCRIPTOR_REVISION))
    {
      ret = GetLastError();
      break;
    }

    // Create a SID for the NT Service\TrustedInstaller group.
    SID_IDENTIFIER_AUTHORITY SIDAuthNT = SECURITY_NT_AUTHORITY;
    if (!AllocateAndInitializeSid(&SIDAuthNT, 6,
      SECURITY_SERVICE_ID_BASE_RID,
      SECURITY_TRUSTED_INSTALLER_RID1,
      SECURITY_TRUSTED_INSTALLER_RID2,
      SECURITY_TRUSTED_INSTALLER_RID3,
      SECURITY_TRUSTED_INSTALLER_RID4,
      SECURITY_TRUSTED_INSTALLER_RID5, 0, 0,
      &psid))
    {
      ret = GetLastError();
      break;
    }

    if (!SetSecurityDescriptorOwner(pNewSD, psid, FALSE))
    {
      ret = GetLastError();
      break;
    }
    ret = RegSetKeySecurity(hk, OWNER_SECURITY_INFORMATION, pNewSD);
    if (ret != ERROR_SUCCESS)
    {
      break;
    }   
  } while (false);
  
  if (pOldSD) LocalFree(pOldSD);
  if (pNewSD) LocalFree(pNewSD);
  if (psid) LocalFree(psid);
  return ret;
}

void EnableTakeOwnershipPrivilege()
{
  HANDLE hToken = NULL;
  LUID luid;
  OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY, &hToken);
  LookupPrivilegeValue(NULL, SE_TAKE_OWNERSHIP_NAME, &luid);
  TOKEN_PRIVILEGES tp;
  tp.PrivilegeCount = 1;
  tp.Privileges[0].Attributes = SE_PRIVILEGE_ENABLED;
  tp.Privileges[0].Luid = luid;
  AdjustTokenPrivileges(hToken, FALSE, &tp, sizeof(TOKEN_PRIVILEGES), NULL, NULL);
  LookupPrivilegeValue(NULL, SE_RESTORE_NAME, &luid);
  tp.Privileges[0].Luid = luid;
  AdjustTokenPrivileges(hToken, FALSE, &tp, sizeof(TOKEN_PRIVILEGES), NULL, NULL);
  CloseHandle(hToken);
}
