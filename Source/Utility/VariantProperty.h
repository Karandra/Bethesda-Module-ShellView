#pragma once
#include "BethesdaModule.hpp"
#include <Propidl.h>

namespace BethesdaModule::ShellView
{
	class VariantProperty final: public PROPVARIANT
	{
		private:
			HRESULT InternalClear();
			void InternalCopy(const PROPVARIANT& source);
			void InternalSetType(VARENUM type);

			void AssignString(std::string_view value);
			void AssignString(std::wstring_view value);

			template<VARENUM type, class TValue1, class TValue2>
			void CreateValue(TValue1& valueStore, TValue2 value)
			{
				valueStore = std::move(value);
				wReserved1 = 0;
				InternalSetType(type);
			}

			template<VARENUM type, class TValue1, class TValue2>
			VariantProperty& AssignValue(TValue1& valueStore, TValue2 value)
			{
				if (vt != type)
				{
					InternalClear();
				}

				valueStore = std::move(value);
				InternalSetType(type);

				return *this;
			}

		public:
			VariantProperty()
			{
				vt = VT_EMPTY;
				wReserved1 = 0;
			}
			VariantProperty(const PROPVARIANT& value)
			{
				vt = VT_EMPTY;
				InternalCopy(value);
			}
			VariantProperty(const VariantProperty& value)
			{
				vt = VT_EMPTY;
				InternalCopy(value);
			}
			VariantProperty(const char* value)
			{
				vt = VT_EMPTY;
				*this = value;
			}
			VariantProperty(const wchar_t* value)
			{
				vt = VT_EMPTY;
				*this = value;
			}
			VariantProperty(const String& value)
			{
				vt = VT_EMPTY;
				*this = value;
			}
			VariantProperty(bool value)
			{
				CreateValue<VT_BOOL>(boolVal, value ? VARIANT_TRUE : VARIANT_FALSE);
			};
			VariantProperty(int8_t value)
			{
				CreateValue<VT_I1>(cVal, value);
			}
			VariantProperty(uint8_t value)
			{
				CreateValue<VT_UI1>(bVal, value);
			}
			VariantProperty(int16_t value)
			{
				CreateValue<VT_I2>(iVal, value);
			}
			VariantProperty(uint16_t value)
			{
				CreateValue<VT_UI2>(uiVal, value);
			}
			VariantProperty(int32_t value)
			{
				CreateValue<VT_I4>(lVal, value);
			}
			VariantProperty(uint32_t value)
			{
				CreateValue<VT_UI4>(ulVal, value);
			}
			VariantProperty(int64_t value)
			{
				CreateValue<VT_I8>(hVal.QuadPart, value);
			}
			VariantProperty(uint64_t value)
			{
				CreateValue<VT_UI8>(uhVal.QuadPart, value);
			}
			VariantProperty(const FILETIME& value)
			{
				CreateValue<VT_FILETIME>(filetime, value);
			}
			~VariantProperty()
			{
				Clear();
			}

		public:
			VARTYPE GetType() const
			{
				return this->vt;
			}
			bool IsEmpty() const
			{
				return GetType() == VT_EMPTY;
			}

			HRESULT Clear();
			HRESULT Copy(const PROPVARIANT& source);
			HRESULT Attach(PROPVARIANT& source);
			HRESULT Detach(PROPVARIANT& destination);

			template<class T = int>
			std::optional<T> ToInteger() const
			{
				switch (vt)
				{
					case VT_I1:
					{
						return static_cast<T>(cVal);
					}
					case VT_UI1:
					{
						return static_cast<T>(bVal);
					}

					case VT_I2:
					{
						return static_cast<T>(iVal);
					}
					case VT_UI2:
					{
						return static_cast<T>(uiVal);
					}

					case VT_I4:
					{
						return static_cast<T>(lVal);
					}
					case VT_UI4:
					{
						return static_cast<T>(ulVal);
					}

					case VT_I8:
					{
						return static_cast<T>(hVal.QuadPart);
					}
					case VT_UI8:
					{
						return static_cast<T>(uhVal.QuadPart);
					}
				};
				return std::nullopt;
			}
			
			std::optional<bool> ToBool() const
			{
				if (vt == VT_BOOL)
				{
					return boolVal != VARIANT_FALSE;
				}
				return std::nullopt;
			}
			std::optional<String> ToString() const
			{
				if (vt == VT_BSTR)
				{
					return bstrVal;
				}
				return std::nullopt;
			}
			std::optional<FILETIME> ToFileTime() const
			{
				if (vt == VT_FILETIME)
				{
					return filetime;
				}
				return std::nullopt;
			}

		public:
			VariantProperty& operator=(const VariantProperty& value)
			{
				InternalCopy(value);
				return *this;
			}
			VariantProperty& operator=(const PROPVARIANT& value)
			{
				InternalCopy(value);
				return *this;
			}
			VariantProperty& operator=(const char* value)
			{
				AssignString(value);
				return *this;
			}
			VariantProperty& operator=(const wchar_t* value)
			{
				AssignString(value);
				return *this;
			}
			VariantProperty& operator=(const String& value)
			{
				AssignString(value.GetView());
				return *this;
			}
			VariantProperty& operator=(bool value)
			{
				return AssignValue<VT_BOOL>(boolVal, value ? VARIANT_TRUE : VARIANT_FALSE);
			}
			VariantProperty& operator=(int8_t value)
			{
				return AssignValue<VT_I1>(cVal, value);
			}
			VariantProperty& operator=(uint8_t value)
			{
				return AssignValue<VT_UI1>(bVal, value);
			}
			VariantProperty& operator=(int16_t value)
			{
				return AssignValue<VT_I2>(iVal, value);
			}
			VariantProperty& operator=(uint16_t value)
			{
				return AssignValue<VT_UI2>(uiVal, value);
			}
			VariantProperty& operator=(int32_t value)
			{
				return AssignValue<VT_I4>(lVal, value);
			}
			VariantProperty& operator=(uint32_t value)
			{
				return AssignValue<VT_UI4>(ulVal, value);
			}
			VariantProperty& operator=(int64_t value)
			{
				return AssignValue<VT_I8>(hVal.QuadPart, value);
			}
			VariantProperty& operator=(uint64_t value)
			{
				return AssignValue<VT_UI8>(uhVal.QuadPart, value);
			}
			VariantProperty& operator=(const FILETIME& value)
			{
				return AssignValue<VT_FILETIME>(filetime, value);
			}

			int Compare(const VariantProperty& other) const;
			bool operator==(const VariantProperty& other) const
			{
				return Compare(other) == 0;
			}
			bool operator!=(const VariantProperty& other) const
			{
				return Compare(other) != 0;
			}

			explicit operator bool() const
			{
				return !IsEmpty();
			}
			bool operator!() const
			{
				return IsEmpty();
			}
	};
}
