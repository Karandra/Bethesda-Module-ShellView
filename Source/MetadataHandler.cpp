#include "stdafx.h"
#include "MetadataHandler.h"
#include "RegisterExtension.h"
#include "DLL.h"
#include "Resources/resource.h"
#include "Utility/PropertyStore.h"
#include "Utility/VariantProperty.h"

namespace
{
	// https://docs.microsoft.com/ru-ru/windows/win32/properties/props-system-author
	// https://github.com/microsoft/Windows-classic-samples/tree/master/Samples/Win7Samples/winui/shell/appshellintegration

	KxFramework::String FormatPropertyNames(const std::vector<KxFramework::String>& names)
	{
		using namespace KxFramework;

		// Reserve with semi-average property name length
		String result;
		result.reserve(names.size() * 24);

		for (const auto& name: names)
		{
			if (result.IsEmpty())
			{
				result += wxS("prop:");
			}
			result += name;
			result += wxS(';');
		}
		return result;
	}
	KxFramework::String StringOrNone(const KxFramework::String& value)
	{
		return !value.IsEmpty() ? value : wxS("<None>");
	}
	KxFramework::String ConcatWithSeparator(const std::vector<KxFramework::String>& items, const KxFramework::String& separator)
	{
		using namespace KxFramework;

		String result;
		for (const auto& item: items)
		{
			if (!result.IsEmpty())
			{
				result += separator;
			}
			result += item;
		}
		return result;
	}
}

namespace BethesdaModule::ShellView
{
	HRESULT MetadataHandler::CreateInstance(REFIID riid, void** ppv)
	{
		return *MakeObjectInstance<MetadataHandler>(riid, ppv);
	}

	HResult MetadataHandler::ReadMorrowind()
	{
		// Seek after to HEDR and skip its record name and following three 32-bit fields
		m_Stream.Seek(16 + 12);

		// These are fixed length
		m_FileInfo.Author = m_Stream.ReadStringACP(32);
		m_FileInfo.Description = m_Stream.ReadStringACP(256);

		// Skip unknown 32-bit value
		m_Stream.Skip<uint32_t>();

		// Read master-files if any
		String recordName = m_Stream.ReadStringASCII(4);
		while (recordName == wxS("MAST"))
		{
			m_FileInfo.RequiredFiles.emplace_back(m_Stream.ReadStringACP(m_Stream.ReadObject<uint32_t>()));

			// Skip DATA 
			m_Stream.Seek(16);
			recordName = m_Stream.ReadStringASCII(4);
		}

		// Morrowind doesn't store anything inside file to help distinguish master from ordinary plugin,
		// so file extension is the only option.
		if (m_Stream.GetFilePath().GetExtension().IsSameAs(wxS("esm"), StringOpFlag::IgnoreCase))
		{
			m_FileInfo.Flags = HeaderFlags::Master;
		}

		m_FileInfo.FormatLevel = FormatLevel::Morrowind;
		return S_OK;
	}
	HResult MetadataHandler::ReadOblivionSkyrim()
	{
		// Skip size
		m_Stream.Skip<uint32_t>();

		// Read flags and seek to CNAM
		m_FileInfo.Flags = FromInt<HeaderFlags>(m_Stream.ReadObject<uint32_t>());

		// Skip (FormID + Version control info)
		m_Stream.Seek(8);

		// See if next 4 bytes are "HEDR". If so it's Oblivion file, if not it's Skyrim.
		const uint32_t temp = m_Stream.ReadObject<uint32_t>();
		if (std::string_view(reinterpret_cast<const char*>(&temp), sizeof(temp)) == "HEDR")
		{
			return ReadOblivion();
		}
		else
		{
			return ReadSkyrim(temp);
		}
	}
	HResult MetadataHandler::ReadOblivion()
	{
		// Skip rest of the HEDR struct
		m_Stream.Seek(14);

		String recordName = m_Stream.ReadStringASCII(4);
		if (recordName == wxS("CNAM"))
		{
			m_FileInfo.Author = m_Stream.ReadStringACP(m_Stream.ReadObject<uint16_t>());
			recordName = m_Stream.ReadStringASCII(4);
		}
		if (recordName == wxS("SNAM"))
		{
			m_FileInfo.Description = m_Stream.ReadStringACP(m_Stream.ReadObject<uint16_t>());
			recordName = m_Stream.ReadStringASCII(4);
		}
		if (recordName == wxS("MAST"))
		{
			do
			{
				m_FileInfo.RequiredFiles.emplace_back(m_Stream.ReadStringACP(m_Stream.ReadObject<uint16_t>()));

				// Skip DATA 
				m_Stream.Seek(14);
				recordName = m_Stream.ReadStringASCII(4);
			}
			while (recordName == wxS("MAST"));
		}

		m_FileInfo.FormatLevel = FormatLevel::Oblivion;
		return S_OK;
	}
	HResult MetadataHandler::ReadSkyrim(uint32_t formVersion)
	{
		// It was the form version
		m_FileInfo.FormVersion = formVersion;

		// Skip HEDR struct
		m_Stream.Seek(18);

		String recordName = m_Stream.ReadStringASCII(4);
		if (recordName == wxS("CNAM"))
		{
			m_FileInfo.Author = m_Stream.ReadStringACP(m_Stream.ReadObject<uint16_t>());
			recordName = m_Stream.ReadStringASCII(4);
		}
		if (recordName == wxS("SNAM"))
		{
			m_FileInfo.Description = m_Stream.ReadStringACP(m_Stream.ReadObject<uint16_t>());
			recordName = m_Stream.ReadStringASCII(4);
		}

		// Read masters
		if (recordName == wxS("MAST"))
		{
			do
			{
				m_FileInfo.RequiredFiles.emplace_back(m_Stream.ReadStringACP(m_Stream.ReadObject<uint16_t>()));

				// Skip DATA 
				m_Stream.Seek(14);
				recordName = m_Stream.ReadStringASCII(4);
			}
			while (recordName == wxS("MAST"));
		}

		m_FileInfo.FormatLevel = FormatLevel::Skyrim;
		return S_OK;
	}

	MetadataHandler::MetadataHandler()
		:m_RefCount(this)
	{
		DllAddRef();
	}
	MetadataHandler::~MetadataHandler()
	{
		DllRelease();
	}

	HRESULT MetadataHandler::QueryInterface(REFIID riid, void** ppv)
	{
		static const QITAB interfaces[] =
		{
			QITABENT(MetadataHandler, IPropertyStore),
			QITABENT(MetadataHandler, IPropertyStoreCapabilities),
			QITABENT(MetadataHandler, IInitializeWithStream),
			{nullptr, 0},
		};
		return QISearch(this, interfaces, riid, ppv);
	}

	HRESULT MetadataHandler::GetCount(DWORD* pcProps)
	{
		*pcProps = 0;
		return S_OK;
	}
	HRESULT MetadataHandler::GetAt(DWORD iProp, PROPERTYKEY* pkey)
	{
		*pkey = PKEY_Null;
		return E_UNEXPECTED;
	}
	HRESULT MetadataHandler::GetValue(REFPROPERTYKEY key, PROPVARIANT* pPropVar)
	{
		if (key == PKEY_Author)
		{
			VariantProperty property;
			property = StringOrNone(m_FileInfo.Author);
			return property.Detach(*pPropVar);
		}
		if (key == PKEY_Comment)
		{
			VariantProperty property;
			property = StringOrNone(m_FileInfo.Description);
			return property.Detach(*pPropVar);
		}
		if (key == PKEY_FileVersion)
		{
			VariantProperty property;
			if (m_FileInfo.FormVersion != 0)
			{
				property = m_FileInfo.FormVersion;
			}
			else
			{
				property = wxS("<Unknown>");
			}
			return property.Detach(*pPropVar);
		}
		if (key == PKEY_ContentType)
		{
			String content = HeaderFlagsDef::ToOrExpression(m_FileInfo.Flags);

			VariantProperty property;
			property = !content.IsEmpty() ? content : wxS("Normal");
			return property.Detach(*pPropVar);
		}
		if (key == PKEY_DataObjectFormat)
		{
			VariantProperty property;
			property = FormatLevelDef::TryToString(m_FileInfo.FormatLevel).value_or(StringView(wxS("<Unknown>")));
			return property.Detach(*pPropVar);
		}
		if (key == PKEY_Keywords)
		{
			// Required files
			VariantProperty property;
			property = StringOrNone(ConcatWithSeparator(m_FileInfo.RequiredFiles, wxS("; ")));
			return property.Detach(*pPropVar);
		}
		return S_FALSE;
	}
	HRESULT MetadataHandler::SetValue(REFPROPERTYKEY key, REFPROPVARIANT propVar)
	{
		// SetValue just updates the internal value cache
		return E_NOTIMPL;
	}
	HRESULT MetadataHandler::Commit()
	{
		// Commit writes the internal value cache back out to the stream passed to Initialize
		return E_NOTIMPL;
	}

	HRESULT STDMETHODCALLTYPE MetadataHandler::IsPropertyWritable(REFPROPERTYKEY key)
	{
		// Properties are writable to allow copy values
		return S_OK;
	}

	HRESULT MetadataHandler::Initialize(IStream* stream, DWORD streamAccess)
	{
		if (m_Stream.Open(*stream) && m_Stream.ReadStringASCII(m_FileInfo.Signature, 4))
		{
			if (m_FileInfo.Signature == wxS("TES3"))
			{
				return *ReadMorrowind();
			}
			else if (m_FileInfo.Signature == wxS("TES4"))
			{
				return *ReadOblivionSkyrim();
			}
			return S_FALSE;
		}
		return *m_Stream.GetLastError();
	}
}

namespace BethesdaModule::ShellView
{
	HResult RegisterMetadataHandler(const MetadataHandlerInfo& info)
	{
		// Register the property handler COM object, and set the options it uses
		RegisterExtension regHandler(UUIDOf<MetadataHandler>(), RegistryBaseKey::LocalMachine);
		HResult hr = regHandler.RegisterInProcServer(info.Description, L"Both");
		if (hr)
		{
			if (hr = regHandler.RegisterInProcServerAttribute(L"ManualSafeSave", TRUE))
			{
				if (hr = regHandler.RegisterInProcServerAttribute(L"EnableShareDenyWrite", TRUE))
				{
					hr = regHandler.RegisterInProcServerAttribute(L"EnableShareDenyNone", TRUE);
				}
			}
		}

		// Property Handler and Kind registrations use a different mechanism than the rest of the file-type association system, and do not use ProgIDs
		if (hr)
		{
			if (hr = regHandler.RegisterPropertyHandler(info.Extension))
			{
				hr = regHandler.RegisterPropertyHandlerOverride(L"System.Kind");
			}
		}

		// Associate our ProgID with the file extension, and write the remainder of the registration data to the ProgID to minimize conflicts with other applications and facilitate easy unregistration
		if (hr)
		{
			if (hr = regHandler.RegisterExtensionWithProgID(info.Extension, info.ProgID))
			{
				if (hr = regHandler.RegisterProgID(info.ProgID, info.Description, IDC_ICON))
				{
					if (!info.NoOpenMessage.IsEmpty())
					{
						hr = regHandler.RegisterProgIDValue(info.ProgID, L"NoOpen", info.NoOpenMessage);
					}

					if (hr)
					{
						//hr = regHandler.RegisterNewMenuNullFile(info.Extension, info.ProgID);

						auto RegisterPropertyNames = [&](const String& type, const std::vector<KxFramework::String>& names) -> HResult
						{
							String formattedNames = FormatPropertyNames(names);
							if (!formattedNames.IsEmpty())
							{
								return regHandler.RegisterProgIDValue(info.ProgID, type, formattedNames);
							}
							return S_OK;
						};

						if (hr = RegisterPropertyNames(L"FullDetails", info.FullDetailsPropertyNames))
						{
							if (hr = RegisterPropertyNames(L"InfoTip", info.InfoTipPropertyNames))
							{
								return RegisterPropertyNames(L"PreviewDetails", info.PreviewDetailsPropertyNames);
							}
						}
					}
				}
			}
		}

		// Also register the property-driven thumbnail handler on the ProgID
		if (hr)
		{
			regHandler.SetHandlerCLSID(UUIDOf<PropertyThumbnailHandler>());
			hr = regHandler.RegisterThumbnailHandler(info.ProgID);
		}
		return hr;
	}
	HResult UnregisterMetadataHandler(const MetadataHandlerInfo& info)
	{
		// Unregister the property handler COM object.
		RegisterExtension regHandler(UUIDOf<MetadataHandler>(), RegistryBaseKey::LocalMachine);
		HResult hr = regHandler.UnRegisterObject();
		if (hr)
		{
			// Unregister the property handler and kind for the file extension.
			if (hr = regHandler.UnRegisterPropertyHandler(info.Extension))
			{
				if (hr = regHandler.UnRegisterKind(info.Extension))
				{
					// Remove the whole ProgID since we own all of those settings.
					// Don't try to remove the file extension association since some other application may have overridden it with their own ProgID in the meantime.
					// Leaving the association to a non-existing ProgID is handled gracefully by the Shell.
					// NOTE: If the file extension is unambiguously owned by this application, the association to the ProgID could be safely removed as well,
					// along with any other association data stored on the file extension itself.
					hr = regHandler.UnRegisterProgID(info.ProgID, info.Extension);
				}
			}
		}
		return hr;
	}
}
