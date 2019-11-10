#pragma once

#include <Windows.h>

#include <ShlObj.h>

#include <string>

class CustomShortcut : public IPersistFile, public IExtractIconW, public IShellLinkW, public IPropertyStore/*, IUnknown*/
{
  ULONG _refCount;
  std::wstring _fileName;

  std::wstring _targetPath;
  std::wstring _iconPath;
public:
  CustomShortcut();
  ~CustomShortcut();

  // IUnknown
  ULONG STDMETHODCALLTYPE AddRef();
  ULONG STDMETHODCALLTYPE Release();
  HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void **ppvObject);

  // IPersist
  HRESULT STDMETHODCALLTYPE GetClassID(CLSID *);

  // IPersistFile
  HRESULT STDMETHODCALLTYPE GetCurFile(LPOLESTR *ppszFileName);
  HRESULT STDMETHODCALLTYPE IsDirty();
  HRESULT STDMETHODCALLTYPE Load(LPCOLESTR pszFileName, DWORD dwMode);
  HRESULT STDMETHODCALLTYPE Save(LPCOLESTR pszFileName, BOOL fRemember);
  HRESULT STDMETHODCALLTYPE SaveCompleted(LPCOLESTR pszFileName);

  // IExtractIconW
  HRESULT STDMETHODCALLTYPE GetIconLocation(UINT uFlags, PWSTR pszIconFile, UINT cchMax, INT *piIndex, UINT *pwFlags);
  HRESULT STDMETHODCALLTYPE Extract(PCWSTR pszFile, UINT nIconIndex, HICON *phiconLarge, HICON *phiconSmall, UINT nIconSize);

  // IShellLinkW
  HRESULT STDMETHODCALLTYPE GetPath(LPWSTR pszFile, int cch, WIN32_FIND_DATAW *pfd, DWORD fFlags);
  HRESULT STDMETHODCALLTYPE GetIDList(PIDLIST_ABSOLUTE *ppidl);
  HRESULT STDMETHODCALLTYPE SetIDList(PCIDLIST_ABSOLUTE pidl);
  HRESULT STDMETHODCALLTYPE GetDescription(LPWSTR pszName, int cch);
  HRESULT STDMETHODCALLTYPE SetDescription(LPCWSTR pszName);
  HRESULT STDMETHODCALLTYPE GetWorkingDirectory(LPWSTR pszDir, int cch);
  HRESULT STDMETHODCALLTYPE SetWorkingDirectory(LPCWSTR pszDir);
  HRESULT STDMETHODCALLTYPE GetArguments(LPWSTR pszArgs, int cch);
  HRESULT STDMETHODCALLTYPE SetArguments(LPCWSTR pszArgs);
  HRESULT STDMETHODCALLTYPE GetHotkey(WORD *pwHotkey);
  HRESULT STDMETHODCALLTYPE SetHotkey(WORD wHotkey);
  HRESULT STDMETHODCALLTYPE GetShowCmd(int *piShowCmd);
  HRESULT STDMETHODCALLTYPE SetShowCmd(int iShowCmd);
  HRESULT STDMETHODCALLTYPE GetIconLocation(LPWSTR pszIconPath, int cch, int *piIcon);
  HRESULT STDMETHODCALLTYPE SetIconLocation(LPCWSTR pszIconPath, int iIcon);
  HRESULT STDMETHODCALLTYPE SetRelativePath(LPCWSTR pszPathRel, DWORD dwReserved);
  HRESULT STDMETHODCALLTYPE Resolve(HWND hwnd, DWORD fFlags);
  HRESULT STDMETHODCALLTYPE SetPath(LPCWSTR pszFile);

  //IPropertyStore
  HRESULT STDMETHODCALLTYPE GetCount(
    /* [out] */ __RPC__out DWORD *cProps);

  HRESULT STDMETHODCALLTYPE GetAt(
    /* [in] */ DWORD iProp,
    /* [out] */ __RPC__out PROPERTYKEY *pkey);

  HRESULT STDMETHODCALLTYPE GetValue(
    /* [in] */ __RPC__in REFPROPERTYKEY key,
    /* [out] */ __RPC__out PROPVARIANT *pv);

  HRESULT STDMETHODCALLTYPE SetValue(
    /* [in] */ __RPC__in REFPROPERTYKEY key,
    /* [in] */ __RPC__in REFPROPVARIANT propvar);

  HRESULT STDMETHODCALLTYPE Commit(void);
};
