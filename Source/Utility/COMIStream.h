#pragma once
#include "BethesdaModule.hpp"
#include <Kx/General/StreamWrappers.h>
#include <Kx/FileSystem/FSPath.h>
#include <shlobj.h>
#include <shlwapi.h>

namespace BethesdaModule::ShellView
{
	class COMIStream:
		public IStreamWrapper,
		public InputStreamWrapper<wxInputStream>,
		public OutputStreamWrapper<wxOutputStream>,
		public IOStreamWrapper<COMIStream>
	{
		private:
			COMPtr<IStream> m_Stream;
			mutable HResult m_LastError = S_OK;

		protected:
			wxFileOffset OnSysTell() const override;
			wxFileOffset OnSysSeek(wxFileOffset pos, wxSeekMode mode) override;
			size_t OnSysRead(void* buffer, size_t size) override;
			size_t OnSysWrite(const void* buffer, size_t size) override;

		public:
			COMIStream() = default;
			COMIStream(IStream& stream)
			{
				Open(stream);
			}

		public:
			bool IsOk() const override
			{
				return m_Stream != nullptr;
			}

			bool Open(IStream& stream)
			{
				m_Stream = nullptr;
				m_LastError = stream.QueryInterface(&m_Stream);

				return m_LastError.IsSuccess();
			}
			bool Close() override
			{
				if (m_Stream)
				{
					m_Stream = nullptr;
					return true;
				}
				return false;
			}
			
			HResult GetLastError() const
			{
				return m_LastError;
			}
			FSPath GetFilePath() const;

			bool IsWriteable() const override;
			bool IsReadable() const override;

			bool Flush() override;
			bool SetAllocationSize(BinarySize offset = {}) override;
	};
}
