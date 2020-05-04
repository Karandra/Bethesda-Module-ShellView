#pragma once
#include "BethesdaModule.hpp"
#include <Kx/System/COM.h>
#include <Kx/System/ErrorCodeValue.h>
#include <propsys.h>

namespace BethesdaModule::ShellView
{
	bool TestForPropertyKey(IPropertyStore* pps, REFPROPERTYKEY pk);

	HResult LoadPropertyStoreFromStream(IStream* pstm, REFIID riid, void** ppv);
	HResult SavePropertyStoreToStream(IPropertyStore* pps, IStream* pstm);
	HResult CopyPropertyStores(IPropertyStore* ppsDest, IPropertyStore* ppsSource);
	HResult ClonePropertyStoreToMemory(IPropertyStore* ppsSource, REFIID riid, void** ppv);
	HResult SavePropertyStoreToString(IPropertyStore* pps, PWSTR* ppszBase64);

	HResult SerializePropVariantAsString(REFPROPVARIANT propvar, PWSTR* ppszOut);
	HResult DeserializePropVariantFromString(PCWSTR pszIn, PROPVARIANT* ppropvar);

	// Use 'LocalFree' to free the result
	HResult GetBase64StringFromStream(IStream* pstm, PWSTR* ppszBase64);
	HResult IStream_ReadToBuffer(IStream* pstm, UINT uMaxSize, BYTE** ppBytes, UINT* pcBytes);
	HResult SaveSerializedPropStorageToStream(IPersistSerializedPropStorage* psps, IStream* pstm);

	HResult GetSafeSaveStream(IStream* pstmIn, IStream** ppstmTarget);
}
