#pragma once
#include "BethesdaModule.hpp"
#include "Utility/COMRefCount.h"
#include <shlwapi.h>
#include <propkey.h>
#include <propsys.h>

#include <Kx/System/COM.h>
#include <Kx/System/ErrorCodeValue.h>
#include <Kx/FileSystem/FSPath.h>

namespace BethesdaModule::ShellView
{
	class MetadataHandler: public IPropertyStore, public IInitializeWithStream
	{
		public:
			static HRESULT CreateInstance(REFIID riid, void** ppv);

		private:
			COMRefCount<MetadataHandler, ULONG, 1> m_RefCount;
			COMPtr<IStream> m_Stream;
			COMPtr<IPropertyStoreCache> m_Cache = nullptr; // Internal value cache to abstract IPropertyStore operations from the DOM back-end
			DWORD m_StreamAccess = 0;

		private:
			HRESULT SaveToStream();

		public:
			MetadataHandler();
			~MetadataHandler();

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

			// IPropertyStore
			HRESULT STDMETHODCALLTYPE GetCount(DWORD* pcProps) override;
			HRESULT STDMETHODCALLTYPE GetAt(DWORD iProp, PROPERTYKEY* pkey) override;
			HRESULT STDMETHODCALLTYPE GetValue(REFPROPERTYKEY key, PROPVARIANT* pPropVar) override;
			HRESULT STDMETHODCALLTYPE SetValue(REFPROPERTYKEY key, REFPROPVARIANT propVar) override;
			HRESULT STDMETHODCALLTYPE Commit() override;

			// IInitializeWithStream
			HRESULT STDMETHODCALLTYPE Initialize(IStream* stream, DWORD streamAccess) override;
	};
}

namespace BethesdaModule::ShellView
{
	HResult RegisterMetadataHandler(const String& extension);
	HResult UnregisterMetadataHandler(const String& extension);
}
