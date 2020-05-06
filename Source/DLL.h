#pragma once
#include <atomic>
#include <windows.h>
#include <shlobj.h>
#include <shlwapi.h>
#include "Utility/COMRefCount.h"

#include "BethesdaModule.hpp"
#include <Kx/System/ErrorCodeValue.h>

namespace BethesdaModule::ShellView
{
	// Our own types
	class DECLSPEC_UUID("F0A39FC0-D248-4EF8-BF47-84A9B44BE002") DLLClassFactory;
	class DECLSPEC_UUID("2A989DB9-CEE7-4B12-B742-51235BCF537B") MetadataHandler;

	// Stuff from the Windows SDK that we will use
	class DECLSPEC_UUID("8d80504a-0826-40c5-97e1-ebc68f953792") DocFilePropertyHandler;
	class DECLSPEC_UUID("97e467b4-98c6-4f19-9588-161b7773d6f6") OfficeDocFilePropertyHandler;
	class DECLSPEC_UUID("45670FA8-ED97-4F44-BC93-305082590BFB") XPSPropertyHandler;
	class DECLSPEC_UUID("c73f6f30-97a0-4ad1-a08f-540d4e9bc7b9") XMLPropertyHandler;
	class DECLSPEC_UUID("993BE281-6695-4BA5-8A2A-7AACBFAAB69E") OfficeOPCPropertyHandler;
	class DECLSPEC_UUID("C41662BB-1FA0-4CE0-8DC5-9B7F8279FF97") OfficeOPCExtractImageHandler;
	class DECLSPEC_UUID("9DBD2C50-62AD-11D0-B806-00C04FD706EC") PropertyThumbnailHandler;
	
	enum class Registration
	{
		Enable,
		Disable
	};
}

namespace BethesdaModule::ShellView
{
	void DllAddRef() noexcept;
	void DllRelease() noexcept;

	HResult RegisterMetatata(Registration registration);
}

namespace BethesdaModule::ShellView
{
	struct ClassObjectInitializer final
	{
		using CreateFunc = HRESULT(*)(REFIID riid, void** ppvObject);

		const CLSID* m_CLSID = nullptr;
		CreateFunc m_CreateFunc = nullptr;
	};

	class DLLClassFactory: public IClassFactory
	{
		public:
			static HRESULT CreateInstance(REFCLSID clsid, const ClassObjectInitializer* classObjectInitializers, size_t initializersCount, REFIID riid, void** result);

		private:
			COMRefCount<DLLClassFactory, ULONG, 1> m_RefCount;
			ClassObjectInitializer::CreateFunc m_CreateFunc = nullptr;

		public:
			DLLClassFactory(ClassObjectInitializer::CreateFunc func)
				:m_RefCount(this), m_CreateFunc(func)
			{
				DllAddRef();
			}
			~DLLClassFactory()
			{
				DllRelease();
			}

		public:
			// IUnknown
			ULONG STDMETHODCALLTYPE AddRef() override
			{
				return m_RefCount.AddRef();
			}
			ULONG STDMETHODCALLTYPE Release() override
			{
				return m_RefCount.Release();
			}
			HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppv) override;

			// IClassFactory
			HRESULT STDMETHODCALLTYPE CreateInstance(IUnknown* punkOuter, REFIID riid, void** ppv) override
			{
				return punkOuter ? CLASS_E_NOAGGREGATION : m_CreateFunc(riid, ppv);
			}
			HRESULT STDMETHODCALLTYPE LockServer(BOOL doLock) override
			{
				if (doLock)
				{
					DllAddRef();
				}
				else
				{
					DllRelease();
				}
				return S_OK;
			}
	};
}

extern "C"
{
	HRESULT STDAPICALLTYPE DllGetClassObject(REFCLSID clsid, REFIID riid, void** result);
	HRESULT STDAPICALLTYPE DllRegisterServer();
	HRESULT STDAPICALLTYPE DllUnregisterServer();
	HRESULT STDAPICALLTYPE DllCanUnloadNow();
}
