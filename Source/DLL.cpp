#include "stdafx.h"
#include "DLL.h"

namespace
{
	std::atomic<size_t> g_RefCount = 0;
}

namespace BethesdaModule::ShellView
{
	void DllAddRef() noexcept
	{
		++g_RefCount;
	}
	void DllRelease() noexcept
	{
		--g_RefCount;
	}
}

namespace BethesdaModule::ShellView
{
	HRESULT DLLClassFactory::CreateInstance(REFCLSID clsid, const ClassObjectInitializer* classObjectInitializers, size_t initializersCount, REFIID riid, void** result)
	{
		if (!classObjectInitializers || initializersCount == 0)
		{
			return CLASS_E_CLASSNOTAVAILABLE;
		}
		if (!result)
		{
			return E_INVALIDARG;
		}
		*result = nullptr;

		HResult hr = CLASS_E_CLASSNOTAVAILABLE;
		for (size_t i = 0; i < initializersCount; i++)
		{
			if (clsid == *classObjectInitializers[i].m_CLSID)
			{
				IClassFactory* classFactory = new(std::nothrow) DLLClassFactory(classObjectInitializers[i].m_CreateFunc);
				if (hr = classFactory ? S_OK : E_OUTOFMEMORY)
				{
					hr = classFactory->QueryInterface(riid, result);
					classFactory->Release();
				}
				break;
			}
		}
		return hr;
	}

	HRESULT STDMETHODCALLTYPE DLLClassFactory::QueryInterface(REFIID riid, void** ppv)
	{
		static const QITAB interfaces[] =
		{
			QITABENT(DLLClassFactory, IClassFactory),
			{nullptr, 0}
		};
		return QISearch(this, interfaces, riid, ppv);
	}
}

extern "C"
{
	HRESULT STDAPICALLTYPE DllGetClassObject(REFCLSID clsid, REFIID riid, void** result)
	{
		using namespace BethesdaModule::ShellView;

		constexpr ClassObjectInitializer classObjectInitializers[] =
		{
			{&__uuidof(COpenMetadataHandler), COpenMetadataHandler_CreateInstance}
		};
		return DLLClassFactory::CreateInstance(clsid, classObjectInitializers, std::size(classObjectInitializers), riid, result);
	}
	HRESULT STDAPICALLTYPE DllRegisterServer()
	{
		using namespace KxFramework;
		using namespace BethesdaModule::ShellView;

		HResult hr = RegisterDocFile();
		if (hr)
		{
			hr = RegisterOpenMetadata();
		}
		return hr;
	}
	HRESULT STDAPICALLTYPE DllUnregisterServer()
	{
		using namespace KxFramework;
		using namespace BethesdaModule::ShellView;

		HResult hr = UnregisterDocFile();
		if (hr)
		{
			hr = UnregisterOpenMetadata();
		}
		return hr;
	}
	HRESULT STDAPICALLTYPE DllCanUnloadNow()
	{
		// Only allow the DLL to be unloaded after all outstanding references have been released
		return g_RefCount == 0 ? S_OK : S_FALSE;
	}

	BOOL STDAPICALLTYPE DllMain(HINSTANCE handle, DWORD eventID, void* reserved)
	{
		return FALSE;
	}
}
