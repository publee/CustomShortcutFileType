#pragma once

#include <Windows.h>

class ClassFactory : public IClassFactory
{
  ULONG _refCount;
  CLSID _class;
public:
  ClassFactory();
  ~ClassFactory();
  bool SetClass(REFCLSID rclsid);

  // IUnknown
  ULONG STDMETHODCALLTYPE AddRef();
  ULONG STDMETHODCALLTYPE Release();
  HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void **ppvObject);

  // IClassFactory
  HRESULT STDMETHODCALLTYPE CreateInstance(IUnknown *pUnkOuter, REFIID riid, void **ppvObject);
  HRESULT STDMETHODCALLTYPE LockServer(BOOL fLock);
};
