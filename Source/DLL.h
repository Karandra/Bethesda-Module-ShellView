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
	void DllAddRef() noexcept;
	void DllRelease() noexcept;

	HResult RegisterDocFile();
	HResult UnregisterDocFile();
	HResult RegisterOpenMetadata();
	HResult UnregisterOpenMetadata();

	class DECLSPEC_UUID("E1A6729E-D430-46a4-A45D-CC042D7F7B36") COpenMetadataHandler;
	HRESULT COpenMetadataHandler_CreateInstance(REFIID riid, void** ppv);
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

BMSV_EXPORT HRESULT STDAPICALLTYPE DllGetClassObject(REFCLSID clsid, REFIID riid, void** result);
BMSV_EXPORT HRESULT STDAPICALLTYPE DllRegisterServer();
BMSV_EXPORT HRESULT STDAPICALLTYPE DllUnregisterServer();
BMSV_EXPORT HRESULT STDAPICALLTYPE DllCanUnloadNow();
