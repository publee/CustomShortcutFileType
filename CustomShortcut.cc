#include "CustomShortcutLibrary.h"
#include "CustomShortcut.h"
#include "Utilities.h"

#include <InitGuid.h>

#include <atlbase.h>

#include <iostream>
#include <codecvt>
#include <fstream>
#include <string>

using namespace std;

//////////////////
// CustomShortcut
////////////////

// {42465C3A-83D3-4310-B27D-F271DE372764}
DEFINE_GUID(CLSID_CustomShortcut,
  0x42465c3a, 0x83d3, 0x4310, 0xb2, 0x7d, 0xf2, 0x71, 0xde, 0x37, 0x27, 0x64);

CustomShortcut::CustomShortcut()
{
  InterlockedIncrement(&g_moduleCount);
  _refCount = 1;
}

CustomShortcut::~CustomShortcut()
{
  InterlockedDecrement(&g_moduleCount);
}

// IUnknown
ULONG STDMETHODCALLTYPE CustomShortcut::AddRef()
{
  return ++_refCount;
}

ULONG STDMETHODCALLTYPE CustomShortcut::Release()
{
  ULONG result = --_refCount;

  if (result == 0)
    delete this;

  return result;
}

HRESULT STDMETHODCALLTYPE CustomShortcut::QueryInterface(REFIID riid, void **ppvObject)
{
  if (riid == IID_IUnknown)
  {
    *ppvObject = this;
    AddRef();
    return S_OK;
  }

  if (riid == IID_IPersistFile)
  {
    *ppvObject = static_cast<IPersistFile *>(this);
    AddRef();
    return S_OK;
  }

  if (riid == IID_IExtractIconW)
  {
    *ppvObject = static_cast<IExtractIconW *>(this);
    AddRef();
    return S_OK;
  }

  if (riid == IID_IShellLinkW)
  {
    *ppvObject = static_cast<IShellLinkW *>(this);
    AddRef();
    return S_OK;
  }

  return E_NOINTERFACE;
}

// IPersist through IPersistFile
HRESULT STDMETHODCALLTYPE CustomShortcut::GetClassID(CLSID *)
{
  return E_NOTIMPL;
}

// IPersistFile
HRESULT STDMETHODCALLTYPE CustomShortcut::GetCurFile(LPOLESTR *ppszFileName)
{
  IMalloc *allocator;

  HRESULT result = CoGetMalloc(1, &allocator);

  if (result != S_OK)
    return result;

  *ppszFileName = (LPOLESTR)allocator->Alloc((_fileName.size() + 1) * 2);

  allocator->Release();

  if (!*ppszFileName)
    return E_OUTOFMEMORY;

  wcscpy_s(*ppszFileName, _fileName.size() + 1, _fileName.c_str());

  return S_OK;
}

HRESULT STDMETHODCALLTYPE CustomShortcut::IsDirty()
{
  // This implementation never modifies the file.
  return S_FALSE;
}

HRESULT STDMETHODCALLTYPE CustomShortcut::Load(LPCOLESTR pszFileName, DWORD dwMode)
{
  try
  {
    _fileName = pszFileName;

    wifstream file(_fileName);

    file.imbue(locale(locale::empty(), new codecvt_utf8<wchar_t, 0x10FFFF, std::consume_header>()));

    if (file)
      getline(file, _targetPath);
    if (file)
      getline(file, _iconPath);

    return S_OK;
  }
  catch (bad_alloc)
  {
    return E_OUTOFMEMORY;
  }
  catch (...)
  {
    return E_FAIL;
  }
}

HRESULT STDMETHODCALLTYPE CustomShortcut::Save(LPCOLESTR pszFileName, BOOL fRemember)
{
  if (pszFileName == NULL)
    return S_OK;
  else
    return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE CustomShortcut::SaveCompleted(LPCOLESTR pszFileName)
{
  return S_OK;
}

// IExtractIconW
HRESULT STDMETHODCALLTYPE CustomShortcut::GetIconLocation(UINT uFlags, PWSTR pszIconFile, UINT cchMax, INT *piIndex, UINT *pwFlags)
{
  try
  {
    wstring iconFileName = _iconPath;

    bool forceDefault = ((uFlags & GIL_DEFAULTICON) == GIL_DEFAULTICON);

    if (forceDefault || !FileExists(iconFileName))
      iconFileName = GetModulePath(g_hInstance);

    if (iconFileName.size() >= cchMax)
      return E_FAIL;

    wcscpy_s(pszIconFile, cchMax, iconFileName.c_str());

    if (piIndex != NULL)
      *piIndex = 0;

    if (pwFlags != NULL)
    {
      *pwFlags = GIL_PERINSTANCE;

      if ((uFlags & GIL_CHECKSHIELD) == GIL_CHECKSHIELD)
        *pwFlags |= GIL_FORCENOSHIELD;
    }

    return S_OK;
  }
  catch (bad_alloc)
  {
    return E_OUTOFMEMORY;
  }
  catch (...)
  {
    return E_FAIL;
  }
}

HRESULT STDMETHODCALLTYPE CustomShortcut::Extract(PCWSTR pszFile, UINT nIconIndex, HICON *phiconLarge, HICON *phiconSmall, UINT nIconSize)
{
  // Let Explorer extract icons.
  return S_FALSE;
}

// IShellLinkW
HRESULT STDMETHODCALLTYPE CustomShortcut::GetPath(LPWSTR pszFile, int cch, WIN32_FIND_DATAW *pfd, DWORD fFlags)
{
  wstring exePath = _targetPath;

  if (!FileExists(exePath))
    exePath = _fileName;

  if ((fFlags & SLGP_SHORTPATH) != 0)
  {
    DWORD shortPathLength = GetShortPathNameW(
      exePath.c_str(),
      pszFile,
      cch);

    if ((shortPathLength > 0) && (shortPathLength < DWORD(cch)))
      return S_OK;
    else
      return E_INVALIDARG;
  }
  else
  {
    wcsncpy_s(pszFile, cch, exePath.c_str(), _TRUNCATE);

    if (exePath.size() < DWORD(cch))
      return S_OK;
    else
      return E_INVALIDARG;
  }
}

HRESULT STDMETHODCALLTYPE CustomShortcut::GetIDList(PIDLIST_ABSOLUTE *ppidl)
{
  CComPtr<IShellFolder> desktop;

  HRESULT result = SHGetDesktopFolder(&desktop);

  if (!SUCCEEDED(result))
    return result;
  else
  {
    wstring exePath = _targetPath;

    if (!FileExists(exePath))
      exePath = _fileName;

    ULONG numberOfCharsEaten;
    ULONG attributes;

    return desktop->ParseDisplayName(
      NULL, // hWnd
      NULL, // pbc (IBindCtx)
      &exePath[0],
      &numberOfCharsEaten,
      ppidl,
      &attributes);
  }
}

HRESULT STDMETHODCALLTYPE CustomShortcut::SetIDList(PCIDLIST_ABSOLUTE pidl)
{
  return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE CustomShortcut::GetDescription(LPWSTR pszName, int cch)
{
  return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE CustomShortcut::SetDescription(LPCWSTR pszName)
{
  return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE CustomShortcut::GetWorkingDirectory(LPWSTR pszDir, int cch)
{
  return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE CustomShortcut::SetWorkingDirectory(LPCWSTR pszDir)
{
  return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE CustomShortcut::GetArguments(LPWSTR pszArgs, int cch)
{
  return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE CustomShortcut::SetArguments(LPCWSTR pszArgs)
{
  return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE CustomShortcut::GetHotkey(WORD *pwHotkey)
{
  *pwHotkey = 0;
  return S_OK;
}

HRESULT STDMETHODCALLTYPE CustomShortcut::SetHotkey(WORD wHotkey)
{
  return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE CustomShortcut::GetShowCmd(int *piShowCmd)
{
  return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE CustomShortcut::SetShowCmd(int iShowCmd)
{
  return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE CustomShortcut::GetIconLocation(LPWSTR pszIconPath, int cch, int *piIcon)
{
  return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE CustomShortcut::SetIconLocation(LPCWSTR pszIconPath, int iIcon)
{
  return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE CustomShortcut::SetRelativePath(LPCWSTR pszPathRel, DWORD dwReserved)
{
  return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE CustomShortcut::Resolve(HWND hwnd, DWORD fFlags)
{
  return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE CustomShortcut::SetPath(LPCWSTR pszFile)
{
  return E_NOTIMPL;
}
