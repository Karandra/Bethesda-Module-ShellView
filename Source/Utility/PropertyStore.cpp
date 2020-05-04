#include "stdafx.h"
#include "PropertyStore.h"
#include <shlobj.h>
#include <propsys.h>
#include <wincrypt.h>
#include <propvarutil.h>

#define MAP_ENTRY(x) {L#x, x}

namespace BethesdaModule::ShellView
{
	bool TestForPropertyKey(IPropertyStore* pps, REFPROPERTYKEY pk)
	{
		PROPVARIANT pv = {};
		bool fHasPropertyKey = SUCCEEDED(pps->GetValue(pk, &pv)) && (pv.vt != VT_EMPTY);
		PropVariantClear(&pv);
		return fHasPropertyKey;
	}

	HResult LoadPropertyStoreFromStream(IStream* pstm, REFIID riid, void** ppv)
	{
		*ppv = nullptr;

		STATSTG stat = {};
		HResult hr = pstm->Stat(&stat, STATFLAG_NONAME);
		if (hr)
		{
			COMPtr<IPersistStream> pps;
			if (hr = PSCreateMemoryPropertyStore(IID_PPV_ARGS(&pps)))
			{
				if (stat.cbSize.QuadPart == 0)
				{
					// Empty stream => empty property store
					hr = S_OK;
				}
				else if (stat.cbSize.QuadPart > (128 * 1024))
				{
					// Can't be that large
					hr = STG_E_MEDIUMFULL;
				}
				else
				{
					hr = pps->Load(pstm);
				}

				if (hr)
				{
					hr = pps->QueryInterface(riid, ppv);
				}
			}
		}
		return hr;
	}
	HResult SavePropertyStoreToStream(IPropertyStore* pps, IStream* pstm)
	{
		// Serializes a property store either directly or by creating a copy and using the memory property store and IPersistSerializedPropStorage

		COMPtr<IPersistSerializedPropStorage> psps;
		HResult hr = pps->QueryInterface(IID_PPV_ARGS(&psps));
		if (hr)
		{
			// Implement size limit here
			hr = SaveSerializedPropStorageToStream(psps, pstm);
		}
		else
		{
			// this store does not support serialization directly, copy it to the one that does and save it
			if (hr = ClonePropertyStoreToMemory(pps, IID_PPV_ARGS(&psps)))
			{
				hr = SaveSerializedPropStorageToStream(psps, pstm);
			}
		}
		return hr;
	}
	HResult CopyPropertyStores(IPropertyStore* ppsDest, IPropertyStore* ppsSource)
	{
		// OK if no properties to copy
		HResult hr = S_OK;

		int iIndex = 0;
		PROPERTYKEY key;
		while (hr && (S_OK == ppsSource->GetAt(iIndex++, &key)))
		{
			PROPVARIANT propvar;
			if (hr = ppsSource->GetValue(key, &propvar))
			{
				hr = ppsDest->SetValue(key, propvar);
				PropVariantClear(&propvar);
			}
		}
		return hr;
	}
	HResult ClonePropertyStoreToMemory(IPropertyStore* ppsSource, REFIID riid, void** ppv)
	{
		// Returns an instance of a memory property store (IPropertyStore) or related interface in the output

		*ppv = nullptr;
		COMPtr<IPropertyStore> ppsMemory;
		HResult hr = PSCreateMemoryPropertyStore(IID_PPV_ARGS(&ppsMemory));
		if (hr)
		{
			if (hr = CopyPropertyStores(ppsMemory, ppsSource))
			{
				hr = ppsMemory->QueryInterface(riid, ppv);
			}
		}
		return hr;
	}
	HResult SavePropertyStoreToString(IPropertyStore* pps, PWSTR* ppszBase64)
	{
		*ppszBase64 = nullptr;

		COMPtr<IStream> pstm;
		HResult hr = ::CreateStreamOnHGlobal(nullptr, TRUE, &pstm);
		if (hr)
		{
			if (hr = SavePropertyStoreToStream(pps, pstm))
			{
				IStream_Reset(pstm);
				hr = GetBase64StringFromStream(pstm, ppszBase64);
			}
		}
		return hr;
	}

	HResult SerializePropVariantAsString(REFPROPVARIANT propvar, PWSTR* ppszOut)
	{
		// Serialize PROPVARIANT to binary form
		ULONG cbBlob = 0;
		SERIALIZEDPROPERTYVALUE* pBlob = nullptr;

		HResult hr = StgSerializePropVariant(&propvar, &pBlob, &cbBlob);
		if (hr)
		{
			// Determine the required buffer size
			hr = E_FAIL;
			DWORD cchString = 0;

			if (::CryptBinaryToStringW((BYTE*)pBlob, cbBlob, CRYPT_STRING_BASE64, nullptr, &cchString))
			{
				// Allocate a sufficient buffer
				hr = E_OUTOFMEMORY;
				*ppszOut = (PWSTR)COM::AllocateMemory(sizeof(wchar_t) * cchString);
				if (*ppszOut)
				{
					// convert the serialized binary blob to a string representation
					hr = E_FAIL;
					if (::CryptBinaryToStringW((BYTE*)pBlob, cbBlob, CRYPT_STRING_BASE64, *ppszOut, &cchString))
					{
						hr = S_OK;
					}
					else
					{
						COM::FreeMemory(*ppszOut);
						*ppszOut = nullptr;
					}
				}
			}
			COM::FreeMemory(pBlob);
		}
		return hr;
	}
	HResult DeserializePropVariantFromString(PCWSTR pszIn, PROPVARIANT* ppropvar)
	{
		HResult hr = E_FAIL;
		DWORD dwFormatUsed = 0;
		DWORD dwSkip = 0;
		DWORD cbBlob = 0;

		// Compute and validate the required buffer size
		if (::CryptStringToBinaryW(pszIn, 0, CRYPT_STRING_BASE64, nullptr, &cbBlob, &dwSkip, &dwFormatUsed) && dwSkip == 0 && dwFormatUsed == CRYPT_STRING_BASE64)
		{
			// Allocate a buffer to hold the serialized binary blob
			hr = E_OUTOFMEMORY;
			if (BYTE* pbSerialized = (BYTE*)COM::AllocateMemory(cbBlob))
			{
				// convert the string to a serialized binary blob
				hr = E_FAIL;
				if (::CryptStringToBinaryW(pszIn, 0, CRYPT_STRING_BASE64, pbSerialized, &cbBlob, &dwSkip, &dwFormatUsed))
				{
					// Deserialize the blob back into a PROPVARIANT value
					hr = ::StgDeserializePropVariant((SERIALIZEDPROPERTYVALUE*)pbSerialized, cbBlob, ppropvar);
				}
				COM::FreeMemory(pbSerialized);
			}
		}
		return hr;
	}

	HResult GetBase64StringFromStream(IStream* pstm, PWSTR* ppszBase64)
	{
		*ppszBase64 = nullptr;

		BYTE* pdata = nullptr;
		UINT cBytes = 0;
		HResult hr = IStream_ReadToBuffer(pstm, 256 * 1024, &pdata, &cBytes);
		if (hr)
		{
			DWORD dwStringSize = 0;
			hr = ::CryptBinaryToStringW(pdata, cBytes, CRYPT_STRING_BASE64, nullptr, &dwStringSize) ? S_OK : E_FAIL;
			if (hr)
			{
				*ppszBase64 = (PWSTR)LocalAlloc(LPTR, dwStringSize * sizeof(**ppszBase64));
				hr = *ppszBase64 ? S_OK : E_OUTOFMEMORY;
				if (hr)
				{
					hr = ::CryptBinaryToStringW(pdata, cBytes, CRYPT_STRING_BASE64, *ppszBase64, &dwStringSize) ? S_OK : E_FAIL;
					if (!hr)
					{
						::LocalFree(*ppszBase64);
						*ppszBase64 = nullptr;
					}
				}
			}
			::LocalFree(pdata);
		}
		return hr;
	}
	HResult IStream_ReadToBuffer(IStream* pstm, UINT uMaxSize, BYTE** ppBytes, UINT* pcBytes)
	{
		*ppBytes = nullptr;
		*pcBytes = 0;

		ULARGE_INTEGER uli = {};
		HResult hr = ::IStream_Size(pstm, &uli);
		if (hr)
		{
			const ULARGE_INTEGER c_uliMaxSize = {uMaxSize};

			hr = (uli.QuadPart < c_uliMaxSize.QuadPart) ? S_OK : E_FAIL;
			if (hr)
			{
				BYTE* pdata = (BYTE*)::LocalAlloc(LPTR, uli.LowPart);
				hr = pdata ? S_OK : E_OUTOFMEMORY;
				if (hr)
				{
					hr = ::IStream_Read(pstm, pdata, uli.LowPart);
					if (hr)
					{
						*ppBytes = pdata;
						*pcBytes = uli.LowPart;
					}
					else
					{
						::LocalFree(pdata);
					}
				}
			}
		}
		return hr;
	}
	HResult SaveSerializedPropStorageToStream(IPersistSerializedPropStorage* psps, IStream* pstm)
	{
		// IPersistSerializedPropStorage implies the standard property store serialization format supported
		// by the memory property store. this function saves the property store into a stream using that format
		// on win7 use IPersistSerializedPropStorage2 methods here.

		COMPtr<IPersistStream> pPersistStream;
		HResult hr = psps->QueryInterface(IID_PPV_ARGS(&pPersistStream));
		if (hr)
		{
			// Implement size limit here
			hr = pPersistStream->Save(pstm, TRUE);
		}
		return hr;
	}
	
	HResult GetSafeSaveStream(IStream* pstmIn, IStream** ppstmTarget)
	{
		// Property handlers should indicate that they support safe safe by providing
		// the "ManualSafeSave" attribute on their CLSID value in the registry.
		// other code that saves files should use this to get safe save functionality
		// when avaiable.

		// Uses IDestinationStreamFactory to discover the stream to save in.
		// after a successful return from this function save to *ppstmTarget
		// and then call Commit(STGC_DEFAULT) on pstmIn.

		*ppstmTarget = nullptr;

		COMPtr<IDestinationStreamFactory> pdsf;
		HResult hr = pstmIn->QueryInterface(&pdsf);
		if (hr)
		{
			hr = pdsf->GetDestinationStream(ppstmTarget);
		}

		if (!hr)
		{
			// Fall back to original stream
			if (hr = pstmIn->QueryInterface(ppstmTarget))
			{
				// We need to reset and truncate it, otherwise we're left with stray bits.
				if (hr = IStream_Reset(*ppstmTarget))
				{
					ULARGE_INTEGER uliNull = {};
					hr = (*ppstmTarget)->SetSize(uliNull);
				}

				if (!hr && (*ppstmTarget))
				{
					(*ppstmTarget)->Release();
				}
			}
		}
		return hr;
	}
}
