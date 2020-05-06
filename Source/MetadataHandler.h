#pragma once
#include "BethesdaModule.hpp"
#include "Utility/COMRefCount.h"
#include "Utility/COMIStream.h"
#include <shlwapi.h>
#include <propkey.h>
#include <propsys.h>

#include <Kx/System/COM.h>
#include <Kx/System/ErrorCodeValue.h>
#include <Kx/General/IndexedEnum.h>
#include <Kx/FileSystem/FSPath.h>

namespace BethesdaModule::ShellView
{
	enum class HeaderFlags: uint32_t
	{
		None = 0,
		Master = 1 << 0,
		Localized = 1 << 7,
		Light = 1 << 9,
		Ignored = 1 << 12,
	};

	struct HeaderFlagsDef final: public IndexedEnumDefinition<HeaderFlagsDef, HeaderFlags, StringView>
	{
		inline static constexpr TItem Items[] =
		{
			{HeaderFlags::Master, wxS("Master")},
			{HeaderFlags::Localized, wxS("Localized")},
			{HeaderFlags::Light, wxS("Light")},
			{HeaderFlags::Ignored, wxS("Ignored")},
		};
	};
}
namespace KxFramework::EnumClass
{
	Kx_EnumClass_AllowEverything(BethesdaModule::ShellView::HeaderFlags);
}

namespace BethesdaModule::ShellView
{
	class MetadataHandler: public IPropertyStore, public IPropertyStoreCapabilities, public IInitializeWithStream
	{
		public:
			static HRESULT CreateInstance(REFIID riid, void** ppv);

		private:
			COMRefCount<MetadataHandler, ULONG, 1> m_RefCount;
			COMIStream m_Stream;

			struct
			{
				String Signature;
				HeaderFlags Flags = HeaderFlags::None;
				uint32_t FormVersion = 0;
				String Author;
				String Description;
			} m_FileInfo;

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

			// IPropertyStoreCapabilities
			HRESULT STDMETHODCALLTYPE IsPropertyWritable(REFPROPERTYKEY key) override;

			// IInitializeWithStream
			HRESULT STDMETHODCALLTYPE Initialize(IStream* stream, DWORD streamAccess) override;
	};
}

namespace BethesdaModule::ShellView
{
	struct MetadataHandlerInfo final
	{
		String Extension;
		String ProgID;
		String Description;
		String NoOpenMessage;

		std::vector<String> FullDetailsPropertyNames;
		std::vector<String> PreviewDetailsPropertyNames;
		std::vector<String> InfoTipPropertyNames;
	};

	HResult RegisterMetadataHandler(const MetadataHandlerInfo& info);
	HResult UnregisterMetadataHandler(const MetadataHandlerInfo& info);
}
