#pragma once
#include <atomic>

namespace BethesdaModule::ShellView
{
	template<class TObject, class T = unsigned long, T initialValue = 0>
	class COMRefCount final
	{
		private:
			TObject* m_Object = nullptr;
			std::atomic<T> m_RefCount = initialValue;

		public:
			COMRefCount(TObject* object) noexcept
				:m_Object(object)
			{
				static_assert(std::is_integral_v<T>, "integer is required");
			}

		public:
			T AddRef() noexcept
			{
				return ++m_RefCount;
			}
			T Release() noexcept
			{
				const T refCount = --m_RefCount;
				if (refCount == 0)
				{
					delete object;
				}
				return refCount;
			}
	};
}
