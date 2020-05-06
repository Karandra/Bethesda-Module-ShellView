#pragma once
#pragma warning(disable: 4005)
#include <Kx/Common.hpp>
#include <Kx/General/String.h>
#include <Kx/General/NativeUUID.h>
#include <Kx/System/COM.h>

namespace BethesdaModule
{
	using namespace KxFramework;

	#undef REFIID
	#define REFIID	const ::IID&

	#undef REFCLSID
	#define REFCLSID	const ::IID&
}

namespace BethesdaModule
{
	template<class T>
	constexpr NativeUUID UUIDOf() noexcept
	{
		const ::GUID guid = __uuidof(T);

		NativeUUID uuid;
		uuid.Data1 = guid.Data1;
		uuid.Data2 = guid.Data2;
		uuid.Data3 = guid.Data3;
		for (size_t i = 0; i < std::size(uuid.Data4); i++)
		{
			uuid.Data4[i] = guid.Data4[i];
		}

		return uuid;
	}

	template<class T, class... Args>
	HResult MakeObjectInstance(REFIID riid, void** ppv, Args&&... arg)
	{
		HResult hr = E_OUTOFMEMORY;
		COMPtr<T> object;

		if (object = new(std::nothrow) T(std::forward<Args>(arg)...))
		{
			hr = object->QueryInterface(riid, ppv);
		}
		return hr;
	}
}
