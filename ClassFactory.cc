#include <Windows.h>

#include "ClassFactory.h"

#include "CustomShortcutLibrary.h"
#include "CustomShortcut.h"

#include <new>

using namespace std;

//////////////////
// ClassFactory
////////////////

ClassFactory::ClassFactory()
{
  InterlockedIncrement(&g_moduleCount);
  _refCount = 1;
}

ClassFactory::~ClassFactory()
{
  InterlockedDecrement(&g_moduleCount);
}

bool ClassFactory::SetClass(REFCLSID rclsid)
{
  if (rclsid == CLSID_CustomShortcut)
  {
    _class = rclsid;
    return true;
  }
  else
    return false;
}

// IUnknown
ULONG STDMETHODCALLTYPE ClassFactory::AddRef()
{
  return ++_refCount;
}

ULONG STDMETHODCALLTYPE ClassFactory::Release()
{
  ULONG result = --_refCount;

  if (result == 0)
    delete this;

  return result;
}

HRESULT STDMETHODCALLTYPE ClassFactory::QueryInterface(REFIID riid, void **ppvObject)
{
  if (riid == IID_IUnknown)
  {
    *ppvObject = this;
    AddRef();
    return S_OK;
  }

  if (riid == IID_IClassFactory)
  {
    *ppvObject = static_cast<IClassFactory *>(this);
    AddRef();
    return S_OK;
  }

  return E_NOINTERFACE;
}

// IClassFactory
HRESULT STDMETHODCALLTYPE ClassFactory::CreateInstance(IUnknown *pUnkOuter, REFIID riid, void **ppvObject)
{
  try
  {
    if (pUnkOuter != NULL)
      return CLASS_E_NOAGGREGATION;

    HRESULT result = CLASS_E_CLASSNOTAVAILABLE;

    if (_class == CLSID_CustomShortcut)
    {
      CustomShortcut *instance = new CustomShortcut();

      result = instance->QueryInterface(riid, ppvObject);

      if (FAILED(result))
        delete instance;
    }

    return result;
  }
  catch (bad_alloc)
  {
    return E_OUTOFMEMORY;
  }
}

HRESULT STDMETHODCALLTYPE ClassFactory::LockServer(BOOL fLock)
{
  // Not used.
  return S_OK;
}
