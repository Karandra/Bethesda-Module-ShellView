#include "stdafx.h"
#include "MetadataHandler.h"
#include "RegisterExtension.h"
#include "DLL.h"
#include "resource.h"
#include "Utility/PropertyStore.h"

namespace
{
	constexpr wchar_t OpenMetadataProgID[] = L"Windows.OpenMetadata";
	constexpr wchar_t OpenMetadataDescription[] = L"Open Metadata File";

	constexpr wchar_t DocFullDetails[] = L"prop:System.PropGroup.Description;System.Title;System.Subject;System.Keywords;System.Category;System.Comment;System.Rating;System.PropGroup.Origin;System.Author;System.Document.LastAuthor;System.Document.RevisionNumber;System.Document.Version;System.ApplicationName;System.Company;System.Document.Manager;System.Document.DateCreated;System.Document.DateSaved;System.Document.DatePrinted;System.Document.TotalEditingTime;System.PropGroup.Content;System.ContentStatus;System.ContentType;System.Document.PageCount;System.Document.WordCount;System.Document.CharacterCount;System.Document.LineCount;System.Document.ParagraphCount;System.Document.Template;System.Document.Scale;System.Document.LinksDirty;System.Language;System.PropGroup.FileSystem;System.ItemNameDisplay;System.ItemType;System.ItemFolderPathDisplay;System.DateCreated;System.DateModified;System.Size;System.FileAttributes;System.OfflineAvailability;System.OfflineStatus;System.SharedWith;System.FileOwner;System.ComputerName";
	constexpr wchar_t DocInfoTip[] = L"prop:System.ItemType;System.Author;System.Size;System.DateModified;System.Document.PageCount";
	constexpr wchar_t DocPreviewDetails[] = L"prop:*System.DateModified;System.Author;System.Keywords;System.Rating;*System.Size;System.Title;System.Comment;System.Category;*System.Document.PageCount;System.ContentStatus;System.ContentType;*System.OfflineAvailability;*System.OfflineStatus;System.Subject;*System.DateCreated;*System.SharedWith";

	KxFramework::NativeUUID FromUUID(const ::UUID& uuid) noexcept
	{
		return *reinterpret_cast<const KxFramework::NativeUUID*>(&uuid);
	}
}

namespace BethesdaModule::ShellView
{
	HRESULT MetadataHandler::CreateInstance(REFIID riid, void** ppv)
	{
		HRESULT hr = E_OUTOFMEMORY;
		if (MetadataHandler* object = new(std::nothrow) MetadataHandler())
		{
			hr = object->QueryInterface(riid, ppv);
			object->Release();
		}
		return hr;
	}

	HRESULT MetadataHandler::SaveToStream()
	{
		COMPtr<IStream> streamSaveTo;
		HResult hr = GetSafeSaveStream(m_Stream, &streamSaveTo);
		if (hr)
		{
			// Write the XML out to the temporary stream and commit it
			if (hr = SavePropertyStoreToStream(m_Cache, streamSaveTo))
			{
				// Also commits 'streamSaveTo'
				hr = m_Stream->Commit(STGC_DEFAULT);
			}
		}
		return *hr;
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
			QITABENT(MetadataHandler, IInitializeWithStream),
			{nullptr, 0},
		};
		return QISearch(this, interfaces, riid, ppv);
	}

	HRESULT MetadataHandler::GetCount(DWORD* pcProps)
	{
		*pcProps = 0;
		return m_Cache ? m_Cache->GetCount(pcProps) : E_UNEXPECTED;
	}
	HRESULT MetadataHandler::GetAt(DWORD iProp, PROPERTYKEY* pkey)
	{
		*pkey = PKEY_Null;
		return m_Cache ? m_Cache->GetAt(iProp, pkey) : E_UNEXPECTED;
	}
	HRESULT MetadataHandler::GetValue(REFPROPERTYKEY key, PROPVARIANT* pPropVar)
	{
		PropVariantInit(pPropVar);
		return m_Cache ? m_Cache->GetValue(key, pPropVar) : E_UNEXPECTED;
	}
	HRESULT MetadataHandler::SetValue(REFPROPERTYKEY key, REFPROPVARIANT propVar)
	{
		// SetValue just updates the internal value cache

		HResult hr = E_UNEXPECTED;
		if (m_Cache)
		{
			// check grfMode to ensure writes are allowed
			hr = STG_E_ACCESSDENIED;
			if (m_StreamAccess & STGM_READWRITE)
			{
				hr = m_Cache->SetValueAndState(key, &propVar, PSC_DIRTY);
			}
		}
		return *hr;
	}
	HRESULT MetadataHandler::Commit()
	{
		// Commit writes the internal value cache back out to the stream passed to Initialize

		HResult hr = E_UNEXPECTED;
		if (m_Cache)
		{
			// Must be opened for writing
			hr = STG_E_ACCESSDENIED;
			if (m_StreamAccess & STGM_READWRITE)
			{
				hr = SaveToStream();
			}
		}
		return *hr;
	}

	HRESULT MetadataHandler::Initialize(IStream* stream, DWORD streamAccess)
	{
		HResult hr = E_UNEXPECTED;
		if (!m_Stream)
		{
			if (hr = LoadPropertyStoreFromStream(stream, IID_PPV_ARGS(&m_Cache)))
			{
				// Save a reference to the stream as well as the 'streamAccess'
				if (hr = stream->QueryInterface(&m_Stream))
				{
					m_StreamAccess = streamAccess;
				}
			}
		}
		return *hr;
	}
}

namespace BethesdaModule::ShellView
{
	HResult RegisterMetadataHandler(const String& extension)
	{
		// Register the property handler COM object, and set the options it uses
		RegisterExtension re(FromUUID(__uuidof(MetadataHandler)), RegistryBaseKey::LocalMachine);
		HResult hr = re.RegisterInProcServer(OpenMetadataDescription, L"Both");
		if (hr)
		{
			if (hr = re.RegisterInProcServerAttribute(L"ManualSafeSave", TRUE))
			{
				if (re.RegisterInProcServerAttribute(L"EnableShareDenyWrite", TRUE))
				{
					hr = re.RegisterInProcServerAttribute(L"EnableShareDenyNone", TRUE);
				}
			}
		}

		// Property Handler and Kind registrations use a different mechanism than the rest of the file-type association system, and do not use ProgIDs
		if (hr)
		{
			if (hr = re.RegisterPropertyHandler(extension))
			{
				hr = re.RegisterPropertyHandlerOverride(L"System.Kind");
			}
		}

		// Associate our ProgID with the file extension, and write the remainder of the registration data to the ProgID to minimize conflicts with other applications and facilitate easy unregistration
		if (hr)
		{
			if (hr = re.RegisterExtensionWithProgID(extension, OpenMetadataProgID))
			{
				if (hr = re.RegisterProgID(OpenMetadataProgID, OpenMetadataDescription, IDC_ICON))
				{
					if (hr = re.RegisterProgIDValue(OpenMetadataProgID, L"NoOpen", L"This is a sample file type and does not have any apps installed to handle it"))
					{
						if (hr = re.RegisterNewMenuNullFile(extension, OpenMetadataProgID))
						{
							if (hr = re.RegisterProgIDValue(OpenMetadataProgID, L"FullDetails", DocFullDetails))
							{
								if (hr = re.RegisterProgIDValue(OpenMetadataProgID, L"InfoTip", DocInfoTip))
								{
									hr = re.RegisterProgIDValue(OpenMetadataProgID, L"PreviewDetails", DocPreviewDetails);
								}
							}
						}
					}
				}
			}
		}

		// Also register the property-driven thumbnail handler on the ProgID
		if (hr)
		{
			re.SetHandlerCLSID(FromUUID(__uuidof(PropertyThumbnailHandler)));
			hr = re.RegisterThumbnailHandler(OpenMetadataProgID);
		}
		return hr;
	}
	HResult UnregisterMetadataHandler(const String& extension)
	{
		// Unregister the property handler COM object.
		RegisterExtension re(FromUUID(__uuidof(MetadataHandler)), RegistryBaseKey::LocalMachine);
		HResult hr = re.UnRegisterObject();
		if (hr)
		{
			// Unregister the property handler and kind for the file extension.
			if (hr = re.UnRegisterPropertyHandler(extension))
			{
				if (hr = re.UnRegisterKind(extension))
				{
					// Remove the whole ProgID since we own all of those settings.
					// Don't try to remove the file extension association since some other application may have overridden it with their own ProgID in the meantime.
					// Leaving the association to a non-existing ProgID is handled gracefully by the Shell.
					// NOTE: If the file extension is unambiguously owned by this application, the association to the ProgID could be safely removed as well,
					// along with any other association data stored on the file extension itself.
					hr = re.UnRegisterProgID(OpenMetadataProgID, extension);
				}
			}
		}
		return hr;
	}
}
