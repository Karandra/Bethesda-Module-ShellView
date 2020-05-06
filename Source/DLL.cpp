#include "stdafx.h"
#include "DLL.h"
#include "MetadataHandler.h"

namespace
{
	std::atomic<size_t> g_RefCount = 0;
	wxAssertHandler_t g_OriginalAssertHandler = nullptr;

	constexpr wchar_t g_MetadataProgID[] = L"BethesdaModule.Metadata";
	const wchar_t* g_Extensions[] =
	{
		L".esp",
		L".esm",
		L".esl",
	};
	const wchar_t* g_Descriptions[] =
	{
		L"Bethesda Module File",
		L"Bethesda Master Module",
		L"Bethesda Light Module",
	};
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

	HResult RegisterMetatata(Registration registration)
	{
		MetadataHandlerInfo info;
		info.ProgID = g_MetadataProgID;
		info.Description = g_Descriptions[0];
		info.FullDetailsPropertyNames =
		{
			wxS("System.ItemNameDisplay"),
			wxS("System.ItemType"),
			wxS("System.ItemFolderPathDisplay"),
			wxS("System.Size"),
			wxS("System.DateCreated"),
			wxS("System.DateModified"),
			wxS("System.FileAttributes"),
			wxS("System.FileOwner"),
			wxS("System.ComputerName"),

			wxS("System.PropGroup.Content"),
			wxS("System.Author"),
			wxS("System.Comment"),
			wxS("System.FileVersion"),
			wxS("System.ContentType"),
			wxS("System.DataObjectFormat"),
			wxS("System.Keywords"),
		};
		info.InfoTipPropertyNames =
		{
			wxS("System.ItemType"),
			wxS("System.Size"),
			wxS("System.DateModified"),

			wxS("System.Author"),
			wxS("System.Comment"),
			wxS("System.FileVersion"),
			wxS("System.ContentType"),
			wxS("System.Keywords")
		};
		info.PreviewDetailsPropertyNames = info.InfoTipPropertyNames;

		HResult hr = S_OK;
		for (size_t i = 0; i < std::size(g_Extensions) && (hr || registration != Registration::Enable); i++)
		{
			info.Extension = g_Extensions[i];
			//info.Description = g_Descriptions[i];

			hr = registration == Registration::Enable ? RegisterMetadataHandler(info) : UnregisterMetadataHandler(info);
		}
		return hr;
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
				COMPtr<IClassFactory> classFactory = new(std::nothrow) DLLClassFactory(classObjectInitializers[i].m_CreateFunc);
				hr = classFactory ? S_OK : E_OUTOFMEMORY;
				if (hr)
				{
					hr = classFactory->QueryInterface(riid, result);
				}
				break;
			}
		}
		return *hr;
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
			{&__uuidof(MetadataHandler), MetadataHandler::CreateInstance}
		};
		return DLLClassFactory::CreateInstance(clsid, classObjectInitializers, std::size(classObjectInitializers), riid, result);
	}
	HRESULT STDAPICALLTYPE DllRegisterServer()
	{
		using namespace BethesdaModule::ShellView;

		return *RegisterMetatata(Registration::Enable);
	}
	HRESULT STDAPICALLTYPE DllUnregisterServer()
	{
		using namespace BethesdaModule::ShellView;

		return *RegisterMetatata(Registration::Disable);
	}
	HRESULT STDAPICALLTYPE DllCanUnloadNow()
	{
		// Only allow the DLL to be unloaded after all outstanding references have been released
		return g_RefCount == 0 ? S_OK : S_FALSE;
	}

	BOOL STDAPICALLTYPE DllMain(HINSTANCE handle, DWORD eventID, void* reserved)
	{
		if (eventID == DLL_PROCESS_ATTACH)
		{
			// Remove all asserts as by default it shows a message dialog which isn't really a good idea for this DLL
			g_OriginalAssertHandler = wxSetAssertHandler(nullptr);
		}
		return TRUE;
	}
}
