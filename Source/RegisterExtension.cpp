#include "stdafx.h"
#include "RegisterExtension.h"
#include <shlobj.h>
#include <shlwapi.h>
#include <strsafe.h>
#include <shobjidl.h>
#pragma comment(lib, "crypt32.lib")
#pragma comment(lib, "shlwapi.lib")

#include <Kx/System/COM.h>
#include <Kx/System/DynamicLibrary.h>
#include <Kx/System/ShellLink.h>
#include <Kx/System/ShellOperations.h>
#include <Kx/General/StringFormater.h>

#pragma warning(disable: 4309)

namespace
{
	EXTERN_C IMAGE_DOS_HEADER __ImageBase;

	KxFramework::HResult ResultFromKnownLastError() noexcept
	{
		using namespace KxFramework;

		const uint32_t errorCode = ::GetLastError();
		return errorCode == ERROR_SUCCESS ? E_FAIL : ::HRESULT_FROM_WIN32(errorCode);
	}
	HMODULE GetThisModuleHandle() noexcept
	{
		// Retrieve the HINSTANCE for the current DLL or EXE using this symbol that
		// the linker provides for every module, avoids the need for a global HINSTANCE variable
		// and provides access to this value for static libraries

		return reinterpret_cast<HMODULE>(&__ImageBase);
	}

	KxFramework::String FormatGUIDToClassID(const KxFramework::UniversallyUniqueID& guid)
	{
		using namespace KxFramework;

		return guid.ToString(UUIDToStringFormat::CurlyBraces|UUIDToStringFormat::UpperCase);
	}
}

namespace BethesdaModule::ShellView
{
	bool RegisterExtension::IsBaseClassProgID(const String& progID) const
	{
		// Work around the missing "NeverDefault" feature for verbs on downlevel platforms
		// these ProgID values should need special treatment to keep the verbs registered there from becoming default.

		return progID.IsSameAs(L"AllFileSystemObjects", StringOpFlag::IgnoreCase) ||
			progID.IsSameAs(L"Directory", StringOpFlag::IgnoreCase) ||
			progID.IsSameAs(L"*", StringOpFlag::IgnoreCase) ||
			progID.IsSameAs(L"SystemFileAssociations\\Directory.", StringOpFlag::IgnoreCase); // SystemFileAssociations\Directory.* values
	}
	HResult RegisterExtension::EnsureBaseProgIDVerbIsNone(const String& progID) const
	{
		// Putting the value of "none" that does not match any of the verbs under this key avoids those verbs from becoming the default.
		if (IsBaseClassProgID(progID))
		{
			RegFormatSetValue(m_RegistryBaseKey, L"Software\\Classes\\%1\\Shell", L"", L"none", progID);
		}
		return S_OK;
	}
	void RegisterExtension::UpdateAssocChanged(HResult hr, const String& keyFormatString) const
	{
		if (hr && !m_AssociationsChanged)
		{
			if (keyFormatString.IsSameAs(wxS("Software\\Classes\\%1"), StringOpFlag::IgnoreCase) ||
				keyFormatString.IsSameAs(wxS("PropertyHandlers"), StringOpFlag::IgnoreCase) ||
				keyFormatString.IsSameAs(wxS("KindMap"), StringOpFlag::IgnoreCase)
				)
			{
				m_AssociationsChanged = true;
			}
		}
	}
	
	HResult RegisterExtension::MapNotFoundToSuccess(HResult hr) const
	{
		return ::HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) == *hr ? HResult(S_OK) : hr;
	}
	HResult RegisterExtension::MapWin32ToHResult(Win32Error errorCode) const
	{
		return ::HRESULT_FROM_WIN32(*errorCode);
	}

	RegisterExtension::RegisterExtension(const UniversallyUniqueID& clsid, RegistryBaseKey hkeyRoot)
		:m_RegistryBaseKey(hkeyRoot)
	{
		SetHandlerCLSID(clsid);
		SetModule(GetThisModuleHandle());
	}
	RegisterExtension::~RegisterExtension()
	{
		if (m_AssociationsChanged)
		{
			// Inform Explorer and such that file association data has changed
			::SHChangeNotify(SHCNE_ASSOCCHANGED, 0, nullptr, nullptr);
		}
	}

	HResult RegisterExtension::SetModule(void* handle)
	{
		if (handle)
		{
			DynamicLibrary library;
			library.AttachHandle(GetThisModuleHandle());
			m_ModuleName = library.GetFilePath();

			if (m_ModuleName.IsValid())
			{
				return S_OK;
			}
		}
		return E_FAIL;
	}

	HResult RegisterExtension::RegisterInProcServer(const String& friendlyName, const String& threadingModel) const
	{
		HResult hr = EnsureModule();
		if (hr)
		{
			String guid = FormatGUIDToClassID(m_ClassGUID);
			if (hr = RegFormatSetValue(m_RegistryBaseKey, L"Software\\Classes\\CLSID\\%1", L"", friendlyName, guid))
			{
				if (hr = RegFormatSetValue(m_RegistryBaseKey, L"Software\\Classes\\CLSID\\%1\\InProcServer32", L"", m_ModuleName, guid))
				{
					hr = RegFormatSetValue(m_RegistryBaseKey, L"Software\\Classes\\CLSID\\%1\\InProcServer32", L"ThreadingModel", threadingModel, guid);
				}
			}
		}
		return hr;
	}
	HResult RegisterExtension::RegisterInProcServerAttribute(const String& attribute, uint32_t value) const
	{
		// Use for:
		// ManualSafeSave = REG_DWORD:<1>
		// EnableShareDenyNone = REG_DWORD:<1>
		// EnableShareDenyWrite = REG_DWORD:<1>

		return RegFormatSetValue(m_RegistryBaseKey, L"Software\\Classes\\CLSID\\%1", attribute, value, FormatGUIDToClassID(m_ClassGUID));
	}
	
	HResult RegisterExtension::RegisterAppAsLocalServer(const String& friendlyName, const String& cmdLine) const
	{
		HResult hr = EnsureModule();
		if (hr)
		{
			String fullCmdLine;
			if (!cmdLine.IsEmpty())
			{
				fullCmdLine = String::Format(wxS("%1 %2"), m_ModuleName.GetFullPathWithNS(), cmdLine);
			}
			else
			{
				fullCmdLine = m_ModuleName.GetFullPathWithNS();
			}

			String guid = FormatGUIDToClassID(m_ClassGUID);
			if (hr = RegFormatSetValue(m_RegistryBaseKey, L"Software\\Classes\\CLSID\\%1\\LocalServer32", L"", fullCmdLine, guid))
			{
				if (hr = RegFormatSetValue(m_RegistryBaseKey, L"Software\\Classes\\CLSID\\%1", L"AppId", guid, guid))
				{
					if (hr = RegFormatSetValue(m_RegistryBaseKey, L"Software\\Classes\\CLSID\\%1", L"", friendlyName, guid))
					{
						// hr = RegFormatSetValue(m_RegistryBaseKey, L"Software\\Classes\\AppID\\%1", L"RunAs", L"Interactive User", guid);
						hr = RegFormatSetValue(m_RegistryBaseKey, L"Software\\Classes\\AppID\\%1", L"", friendlyName, guid);
					}
				}
			}
		}
		return hr;
	}
	HResult RegisterExtension::RegisterElevatableLocalServer(const String& friendlyName, uint32_t idLocalizeString, uint32_t idIconRef) const
	{
		HResult hr = EnsureModule();
		if (hr)
		{
			String guid = FormatGUIDToClassID(m_ClassGUID);
			if (hr = RegFormatSetValue(RegistryBaseKey::LocalMachine, L"Software\\Classes\\CLSID\\%1", L"", friendlyName, guid))
			{
				String resourceString = String::Format(wxS("@%1,-%2"), m_ModuleName.GetFullPathWithNS(), idLocalizeString);
				if (hr = RegFormatSetValue(RegistryBaseKey::LocalMachine, L"Software\\Classes\\CLSID\\%1", L"LocalizedString", resourceString, guid))
				{
					if (hr = RegFormatSetValue(RegistryBaseKey::LocalMachine, L"Software\\Classes\\CLSID\\%1\\LocalServer32", L"", m_ModuleName.GetFullPathWithNS(), guid))
					{
						hr = RegFormatSetValue(RegistryBaseKey::LocalMachine, L"Software\\Classes\\CLSID\\%1\\Elevation", L"Enabled", 1, guid);
						if (idIconRef != 0)
						{
							resourceString = String::Format(wxS("@%1,-%2"), m_ModuleName.GetFullPathWithNS(), idIconRef);
							hr = RegFormatSetValue(RegistryBaseKey::LocalMachine, L"Software\\Classes\\CLSID\\%1\\Elevation", L"IconReference", resourceString, guid);
						}
					}
				}
			}
		}
		return hr;
	}
	HResult RegisterExtension::RegisterElevatableInProcServer(const String& friendlyName, uint32_t idLocalizeString, uint32_t idIconRef) const
	{
		// Shell32\Shell32.man uses this for InProcServer32 cases
		// 010004805800000068000000000000001400000002004400030000000000140003000000010100000000000504000000000014000700000001010000000000050a00000000001400030000000101000000000005120000000102000000000005200000002002000001020000000000052000000020020000
		constexpr char accessPermission[] =
		{
			0x01, 0x00, 0x04, 0x80, 0x60, 0x00, 0x00, 0x00, 0x70, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x14,
			0x00, 0x00, 0x00, 0x02, 0x00, 0x4c, 0x00, 0x03, 0x00, 0x00, 0x00, 0x00, 0x00, 0x14, 0x00, 0x03, 0x00,
			0x00, 0x00, 0x01, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x05, 0x12, 0x00, 0x00, 0x00, 0x00, 0x00, 0x14,
			0x00, 0x07, 0x00, 0x00, 0x00, 0x01, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x05, 0x0a, 0x00, 0x00, 0x00,
			0x00, 0x00, 0x14, 0x00, 0x03, 0x00, 0x00, 0x00, 0x01, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x05, 0x04,
			0x00, 0x00, 0x00, 0xcd, 0xcd, 0xcd, 0xcd, 0xcd, 0xcd, 0xcd, 0xcd, 0x01, 0x02, 0x00, 0x00, 0x00, 0x00,
			0x00, 0x05, 0x20, 0x00, 0x00, 0x00, 0x20, 0x02, 0x00, 0x00, 0x01, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00,
			0x05, 0x20, 0x00, 0x00, 0x00, 0x20, 0x02, 0x00, 0x00
		};

		// And what is this?
		constexpr char launchPermission[] =
		{
			0x01, 0x00, 0x04, 0x80, 0x78, 0x00, 0x00, 0x00, 0x88, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x14,
			0x00, 0x00, 0x00, 0x02, 0x00, 0x64, 0x00, 0x04, 0x00, 0x00, 0x00, 0x00, 0x00, 0x14, 0x00, 0x1f, 0x00,
			0x00, 0x00, 0x01, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x05, 0x12, 0x00, 0x00, 0x00, 0x00, 0x00, 0x18,
			0x00, 0x1f, 0x00, 0x00, 0x00, 0x01, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0x05, 0x20, 0x00, 0x00, 0x00,
			0x20, 0x02, 0x00, 0x00, 0x00, 0x00, 0x14, 0x00, 0x1f, 0x00, 0x00, 0x00, 0x01, 0x01, 0x00, 0x00, 0x00,
			0x00, 0x00, 0x05, 0x04, 0x00, 0x00, 0x00, 0x00, 0x00, 0x14, 0x00, 0x0b, 0x00, 0x00, 0x00, 0x01, 0x01,
			0x00, 0x00, 0x00, 0x00, 0x00, 0x05, 0x12, 0x00, 0x00, 0x00, 0xcd, 0xcd, 0xcd, 0xcd, 0xcd, 0xcd, 0xcd,
			0xcd, 0x01, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0x05, 0x20, 0x00, 0x00, 0x00, 0x20, 0x02, 0x00, 0x00,
			0x01, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0x05, 0x20, 0x00, 0x00, 0x00, 0x20, 0x02, 0x00, 0x00
		};

		HResult hr = EnsureModule();
		if (hr)
		{
			String guid = FormatGUIDToClassID(m_ClassGUID);
			if (hr = RegFormatSetValue(RegistryBaseKey::LocalMachine, L"Software\\Classes\\AppId\\%1", L"", friendlyName, guid))
			{
				hr = RegFormatSetValue(RegistryBaseKey::LocalMachine, L"Software\\Classes\\AppId\\%1", L"AccessPermission", wxScopedCharBuffer::CreateNonOwned(accessPermission, std::size(accessPermission)), guid);
				hr = RegFormatSetValue(RegistryBaseKey::LocalMachine, L"Software\\Classes\\AppId\\%1", L"LaunchPermission", wxScopedCharBuffer::CreateNonOwned(launchPermission, std::size(launchPermission)), guid);

				if (hr = RegFormatSetValue(RegistryBaseKey::LocalMachine, L"Software\\Classes\\CLSID\\%1", L"", friendlyName, guid))
				{
					if (hr = RegFormatSetValue(RegistryBaseKey::LocalMachine, L"Software\\Classes\\CLSID\\%1", L"AppId", guid, guid))
					{
						String resourceString = String::Format(wxS("@%1,-%2"), m_ModuleName.GetFullPathWithNS(), idLocalizeString);
						if (hr = RegFormatSetValue(RegistryBaseKey::LocalMachine, L"Software\\Classes\\CLSID\\%1", L"LocalizedString", resourceString, guid))
						{
							if (hr = RegFormatSetValue(RegistryBaseKey::LocalMachine, L"Software\\Classes\\CLSID\\%1\\InProcServer32", L"", m_ModuleName.GetFullPathWithNS(), guid))
							{
								hr = RegFormatSetValue(RegistryBaseKey::LocalMachine, L"Software\\Classes\\CLSID\\%1\\Elevation", L"Enabled", 1, guid);
								if (hr && idIconRef)
								{
									resourceString = String::Format(wxS("@%1,-%2"), m_ModuleName.GetFullPathWithNS(), idIconRef);
									hr = RegFormatSetValue(RegistryBaseKey::LocalMachine, L"Software\\Classes\\CLSID\\%1\\Elevation", L"IconReference", resourceString, guid);
								}
							}
						}
					}
				}
			}
		}
		return hr;
	}

	HResult RegisterExtension::UnRegisterObject() const
	{
		HResult hr = E_FAIL;
		String guid = FormatGUIDToClassID(m_ClassGUID);

		// Might have an AppID value, trying that.
		if (hr = RegFormatRemoveKey(m_RegistryBaseKey, L"Software\\Classes\\AppID\\%1", guid))
		{
			hr = RegFormatRemoveKey(m_RegistryBaseKey, L"Software\\Classes\\CLSID\\%1", guid);
		}
		return hr;
	}
	HResult RegisterExtension::RegisterAppDropTarget() const
	{
		HResult hr = EnsureModule();
		if (hr)
		{
			// Windows 7 supports per user AppPaths, downlevel requires HKLM
			String guid = FormatGUIDToClassID(m_ClassGUID);
			hr = RegFormatSetValue(m_RegistryBaseKey, L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\%1", L"DropTarget", guid, m_ModuleName.GetName());
		}
		return hr;
	}

	HResult RegisterExtension::RegisterCreateProcessVerb(const String& progID, const String& verb, const String& cmdLine, const String& verbDisplayName) const
	{
		// Make sure no existing registration exists, ignore failure.
		UnRegisterVerb(progID, verb);

		HResult hr = RegFormatSetValue(m_RegistryBaseKey, L"Software\\Classes\\%1\\shell\\%2\\command", L"", cmdLine, progID, verb);
		if (hr)
		{
			hr = EnsureBaseProgIDVerbIsNone(progID);
			hr = RegFormatSetValue(m_RegistryBaseKey, L"Software\\Classes\\%1\\shell\\%2", L"", verbDisplayName, progID, verb);
		}
		return hr;
	}
	HResult RegisterExtension::RegisterDropTargetVerb(const String& progID, const String& verb, const String& verbDisplayName) const
	{
		// Make sure no existing registration exists, ignore failure
		UnRegisterVerb(progID, verb);

		HResult hr = RegFormatSetValue(m_RegistryBaseKey, L"Software\\Classes\\%1\\Shell\\%2\\DropTarget", L"CLSID", FormatGUIDToClassID(m_ClassGUID), progID, verb);
		if (hr)
		{
			hr = EnsureBaseProgIDVerbIsNone(progID);
			hr = RegFormatSetValue(m_RegistryBaseKey, L"Software\\Classes\\%1\\Shell\\%2", L"", verbDisplayName, progID, verb);
		}
		return hr;
	}
	HResult RegisterExtension::RegisterExecuteCommandVerb(const String& progID, const String& verb, const String& verbDisplayName) const
	{
		// Make sure no existing registration exists, ignore failure
		UnRegisterVerb(progID, verb);

		HResult hr = RegFormatSetValue(m_RegistryBaseKey, L"Software\\Classes\\%1\\Shell\\%2\\command", L"DelegateExecute", FormatGUIDToClassID(m_ClassGUID), progID, verb);
		if (hr)
		{
			hr = EnsureBaseProgIDVerbIsNone(progID);
			hr = RegFormatSetValue(m_RegistryBaseKey, L"Software\\Classes\\%1\\Shell\\%2", L"", verbDisplayName, progID, verb);
		}
		return hr;
	}
	HResult RegisterExtension::RegisterExplorerCommandVerb(const String& progID, const String& verb, const String& verbDisplayName) const
	{
		// Must be an InProc handler registered here.
		// Make sure no existing registration exists, ignore failure.
		UnRegisterVerb(progID, verb);

		HResult hr = RegFormatSetValue(m_RegistryBaseKey, L"Software\\Classes\\%1\\Shell\\%2", L"ExplorerCommandHandler", FormatGUIDToClassID(m_ClassGUID), progID, verb);
		if (hr)
		{
			hr = EnsureBaseProgIDVerbIsNone(progID);
			hr = RegFormatSetValue(m_RegistryBaseKey, L"Software\\Classes\\%1\\Shell\\%2", L"", verbDisplayName, progID, verb);
		}
		return hr;
	}
	HResult RegisterExtension::RegisterExplorerCommandStateHandler(const String& progID, const String& verb) const
	{
		return RegFormatSetValue(m_RegistryBaseKey, L"Software\\Classes\\%1\\Shell\\%2", L"CommandStateHandler", FormatGUIDToClassID(m_ClassGUID), progID, verb);
	}
	HResult RegisterExtension::RegisterVerbAttribute(const String& progID, const String& verb, const String& valueName) const
	{
		// NeverDefault
		// LegacyDisable
		// Extended
		// OnlyInBrowserWindow
		// ProgrammaticAccessOnly
		// SeparatorBefore
		// SeparatorAfter
		// CheckSupportedTypes, used SupportedTypes that is a file type filter registered under AppPaths (I think)
		return RegFormatSetValue(m_RegistryBaseKey, L"Software\\Classes\\%1\\shell\\%2", valueName, L"", progID, verb);
	}
	HResult RegisterExtension::RegisterVerbAttribute(const String& progID, const String& verb, const String& valueName, const String& value) const
	{
		// MUIVerb=@dll,-resid
		// MultiSelectModel=Single|Player|Document
		// Position=Bottom|Top
		// DefaultAppliesTo=System.ItemName:"foo"
		// HasLUAShield=System.ItemName:"bar"
		// AppliesTo=System.ItemName:"foo"

		return RegFormatSetValue(m_RegistryBaseKey, L"Software\\Classes\\%1\\shell\\%2", valueName, value, progID, verb);
	}
	HResult RegisterExtension::RegisterVerbAttribute(const String& progID, const String& verb, const String& valueName, uint32_t value) const
	{
		// BrowserFlags
		// ExplorerFlags
		// AttributeMask
		// AttributeValue
		// ImpliedSelectionModel
		// SuppressionPolicy
		return RegFormatSetValue(m_RegistryBaseKey, L"Software\\Classes\\%1\\shell\\%2", valueName, value, progID, verb);
	}
	HResult RegisterExtension::RegisterVerbDefaultAndOrder(const String& progID, const String& verbOrderFirstIsDefault) const
	{
		// "open explorer" is an example
		return RegFormatSetValue(m_RegistryBaseKey, L"Software\\Classes\\%1\\Shell", L"", verbOrderFirstIsDefault, progID);
	}
	HResult RegisterExtension::RegisterPlayerVerbs(const std::vector<String> associations, const String& verb, const String& title) const
	{
		// Register a verb on an array of ProgIDs

		HResult hr = RegisterAppAsLocalServer(title);
		if (hr)
		{
			// Enable this handler to work with OpenSearch results, avoiding the download and open behavior by indicating that we can accept all URL forms
			hr = RegisterHandlerSupportedProtocols(L"*");

			for (const String& value: associations)
			{
				hr = RegisterExecuteCommandVerb(value, verb, title);
				if (hr)
				{
					hr = RegisterVerbAttribute(value, verb, L"NeverDefault");
					if (hr)
					{
						hr = RegisterVerbAttribute(value, verb, L"MultiSelectModel", L"Player");
					}
				}

				if (!hr)
				{
					break;
				}
			}
		}
		return hr;
	}

	HResult RegisterExtension::UnRegisterVerb(const String& progID, const String& verb) const
	{
		return RegFormatRemoveKey(m_RegistryBaseKey, L"Software\\Classes\\%1\\Shell\\%2", progID, verb);
	}
	HResult RegisterExtension::UnRegisterVerbs(const std::vector<String> associations, const String& verb) const
	{
		HResult hr = S_OK;
		for (const String& value: associations)
		{
			hr = UnRegisterVerb(value, verb);
			if (!hr)
			{
				break;
			}
		}

		if (hr && HasClassID())
		{
			hr = UnRegisterObject();
		}
		return hr;
	}

	HResult RegisterExtension::RegisterContextMenuHandler(const String& progID, const String& description) const
	{
		// In process context menu handler
		return RegFormatSetValue(m_RegistryBaseKey, L"Software\\Classes\\%1\\shellex\\ContextMenuHandlers\\%2", L"", description, progID, FormatGUIDToClassID(m_ClassGUID));
	}
	HResult RegisterExtension::RegisterRightDragContextMenuHandler(const String& progID, const String& description) const
	{
		// In process context menu handler for right drag context menu need to create new method that allows out of process handling of this.
		// progID "Folder" or "Directory"
		return RegFormatSetValue(m_RegistryBaseKey, L"Software\\Classes\\%1\\shellex\\DragDropHandlers\\%2", L"", description, progID, FormatGUIDToClassID(m_ClassGUID));
	}

	HResult RegisterExtension::RegisterAppShortcutInSendTo() const
	{
		FSPath appPath = DynamicLibrary::GetCurrentModule().GetFilePath();
		if (!appPath)
		{
			return ResultFromKnownLastError();
		}

		KxFramework::ShellLink shellLink;
		if (shellLink)
		{
			HResult hr = shellLink.SetTarget(appPath);
			if (hr)
			{
				FSPath linkPath = KxFramework::Shell::GetKnownDirectory(KnownDirectoryID::SendTo);
				if (linkPath)
				{
					linkPath /= appPath.GetName();
					linkPath.SetExtension(wxS("lnk"));

					return shellLink.Save(linkPath);
				}
			}
			return hr;
		}
		return E_FAIL;
	}
	HResult RegisterExtension::RegisterThumbnailHandler(const String& extension) const
	{
		// IThumbnailHandler
		// HKEY_CLASSES_ROOT\.wma\ShellEx\{e357fccd-a995-4576-b01f-234630154e96}={9DBD2C50-62AD-11D0-B806-00C04FD706EC}

		return RegFormatSetValue(m_RegistryBaseKey, L"Software\\Classes\\%1\\ShellEx\\{e357fccd-a995-4576-b01f-234630154e96}", L"", FormatGUIDToClassID(m_ClassGUID), extension);
	}
	HResult RegisterExtension::RegisterPropertyHandler(const String& extension) const
	{
		// IPropertyHandler
		// HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\PropertySystem\PropertyHandlers\.docx={993BE281-6695-4BA5-8A2A-7AACBFAAB69E}

		return RegFormatSetValue(RegistryBaseKey::LocalMachine, L"Software\\Microsoft\\Windows\\CurrentVersion\\PropertySystem\\PropertyHandlers\\%1", L"", FormatGUIDToClassID(m_ClassGUID), extension);
	}
	HResult RegisterExtension::RegisterLinkHandler(const String& progID) const
	{
		// IResolveShellLink handler, used for custom link resolution behavior
		return RegFormatSetValue(m_RegistryBaseKey, L"Software\\Classes\\%1\\ShellEx\\LinkHandler", L"", FormatGUIDToClassID(m_ClassGUID), progID);
	}
	HResult RegisterExtension::UnRegisterPropertyHandler(const String& extension) const
	{
		return RegFormatRemoveKey(RegistryBaseKey::LocalMachine, L"Software\\Microsoft\\Windows\\CurrentVersion\\PropertySystem\\PropertyHandlers\\%1", extension);
	}

	HResult RegisterExtension::RegisterProgID(const String& progID, const String& typeName, uint32_t idIcon) const
	{
		// HKCR\<ProgID> = <Type Name>
		//      DefaultIcon=<icon ref>
		//      <icon ref>=<module path>,<res_id>

		HResult hr = RegFormatSetValue(m_RegistryBaseKey, L"Software\\Classes\\%1", L"", typeName, progID);
		if (hr)
		{
			if (idIcon != 0)
			{
				// HKCR\<ProgID>\DefaultIcon
				String iconRef = String::Format(wxS("\"%1\",-%2"), m_ModuleName.GetFullPathWithNS(), idIcon);
				hr = RegFormatSetValue(m_RegistryBaseKey, L"Software\\Classes\\%1\\DefaultIcon", L"", iconRef, progID);
			}
		}
		return hr;
	}
	HResult RegisterExtension::UnRegisterProgID(const String& progID, const String& fileExtension) const
	{
		HResult hr = RegFormatRemoveKey(m_RegistryBaseKey, L"Software\\Classes\\%1", progID);
		if (hr && !fileExtension.IsEmpty())
		{
			hr = RegFormatRemoveKey(m_RegistryBaseKey, L"Software\\Classes\\%1\\%2", fileExtension, progID);
		}
		return hr;
	}
	HResult RegisterExtension::RegisterExtensionWithProgID(const String& fileExtension, const String& progID) const
	{
		// This is where the file association is being taken over

		// HKCR\<.ext>=<ProgID>
		// "Content Type"
		// "PerceivedType"

		// TODO: to be polite if there is an existing mapping of extension to ProgID make sure it is
		// added to the OpenWith list so that users can get back to the old app using OpenWith
		// TODO: verify that HKLM/HKCU settings do not already exist as if they do they will
		// get in the way of the setting being made here
		return RegFormatSetValue(m_RegistryBaseKey, L"Software\\Classes\\%1", L"", progID, fileExtension);
	}
	HResult RegisterExtension::RegisterOpenWith(const String& fileExtension, const String& progID) const
	{
		// Adds the ProgID to a file extension assuming that this ProgID will have the "open" verb under it that will be used in Open With
		return RegFormatSetValue(m_RegistryBaseKey, L"Software\\Classes\\%1\\OpenWithProgIds", progID, L"", fileExtension);
	}
	HResult RegisterExtension::RegisterNewMenuNullFile(const String& fileExtension, const String& progID) const
	{
		// There are 2 forms of this
		// HKCR\<.ext>\ShellNew
		// HKCR\<.ext>\ShellNew\<ProgID> - only
		// ItemName
		// NullFile
		// Data - REG_BINARY:<binary data>
		// File
		// command
		// iconpath

		// Another way that works
		// HKEY_CLASSES_ROOT\.doc\Word.Document.8\ShellNew
		if (!progID.IsEmpty())
		{
			return RegFormatSetValue(m_RegistryBaseKey, L"Software\\Classes\\%1\\%2\\ShellNew", L"NullFile", L"", fileExtension, progID);
		}
		else
		{
			return RegFormatSetValue(m_RegistryBaseKey, L"Software\\Classes\\%1\\ShellNew", L"NullFile", L"", fileExtension);
		}
	}
	HResult RegisterExtension::RegisterNewMenuData(const String& fileExtension, const String& progID, PCSTR pszBase64) const
	{
		if (!progID.IsEmpty())
		{
			return RegFormatSetValue(m_RegistryBaseKey, L"Software\\Classes\\%1\\%2\\ShellNew", L"Data", wxScopedCharBuffer::CreateNonOwned(pszBase64), fileExtension, progID);
		}
		else
		{
			return RegFormatSetValue(m_RegistryBaseKey, L"Software\\Classes\\%1\\ShellNew", L"Data", wxScopedCharBuffer::CreateNonOwned(pszBase64), fileExtension);
		}
	}
	HResult RegisterExtension::RegisterKind(const String& fileExtension, PCWSTR pszKindValue) const
	{
		// Define the kind of a file extension. This is a multi-value property, see 'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\KindMap'
		return RegFormatSetValue(RegistryBaseKey::LocalMachine, L"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\KindMap", fileExtension, pszKindValue);
	}
	HResult RegisterExtension::UnRegisterKind(const String& fileExtension) const
	{
		return RegFormatRemoveValue(RegistryBaseKey::LocalMachine, L"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\KindMap", fileExtension);
	}
	
	HResult RegisterExtension::RegisterPropertyHandlerOverride(PCWSTR pszProperty) const
	{
		// When indexing it is possible to override some of the file system property values, that includes the following
		// use this registration helper to set the override flag for each:
		// System.ItemNameDisplay
		// System.SFGAOFlags
		// System.Kind
		// System.FileName
		// System.ItemPathDisplay
		// System.ItemPathDisplayNarrow
		// System.ItemFolderNameDisplay
		// System.ItemFolderPathDisplay
		// System.ItemFolderPathDisplayNarrow

		return RegFormatSetValue(m_RegistryBaseKey, L"Software\\Classes\\CLSID\\%1\\OverrideFileSystemProperties", pszProperty, 1, FormatGUIDToClassID(m_ClassGUID));
	}
	HResult RegisterExtension::RegisterHandlerSupportedProtocols(const String& protocol) const
	{
		// Protocol values:
		// "*" - all
		// "http"
		// "ftp"
		// "shellstream" - NYI in Win7 (WTF is 'NYI'!?)

		return RegFormatSetValue(m_RegistryBaseKey, L"Software\\Classes\\CLSID\\%1\\SupportedProtocols", protocol, L"", FormatGUIDToClassID(m_ClassGUID));
	}

	HResult RegisterExtension::RegisterProgIDValue(const String& progID, const String& valueName) const
	{
		// Value names that do not require a value
		// HKCR\<ProgID>
		//      NoOpen - display the "No Open" dialog for this file to disable double click
		//      IsShortcut - report SFGAO_LINK for this item type, should have a IShellLink handler
		//      NeverShowExt - never show the file extension
		//      AlwaysShowExt - always show the file extension
		//      NoPreviousVersions - don't display the "Previous Versions" verb for this file type

		return RegFormatSetValue(m_RegistryBaseKey, L"Software\\Classes\\%1", valueName, L"", progID);
	}
	HResult RegisterExtension::RegisterProgIDValue(const String& progID, const String& valueName, const String& value) const
	{
		// value names that require a string value
		// HKCR\<ProgID>
		//      NoOpen - display the "No Open" dialog for this file to disable double click, display this message
		//      FriendlyTypeName - localized resource
		//      ConflictPrompt
		//      FullDetails
		//      InfoTip
		//      QuickTip
		//      PreviewDetails
		//      PreviewTitle
		//      TileInfo
		//      ExtendedTileInfo
		//      SetDefaultsFor - right click.new will populate the file with these properties, example: "prop:System.Author;System.Document.DateCreated"

		return RegFormatSetValue(m_RegistryBaseKey, L"Software\\Classes\\%1", valueName, value, progID);
	}
	HResult RegisterExtension::RegisterProgIDValue(const String& progID, const String& valueName, uint32_t value) const
	{
		// value names that require a uint32_t value
		// HKCR\<ProgID>
		//      EditFlags
		//      ThumbnailCutoff
		return RegFormatSetValue(m_RegistryBaseKey, L"Software\\Classes\\%1", valueName, value, progID);
	}
}
