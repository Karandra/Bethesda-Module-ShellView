#include "stdafx.h"
#include "COMIStream.h"

namespace BethesdaModule::ShellView
{
	wxFileOffset COMIStream::OnSysTell() const
	{
		if (m_Stream)
		{
			LARGE_INTEGER move = {};
			ULARGE_INTEGER newPos = {};
			if (m_LastError = m_Stream->Seek(move, STREAM_SEEK_CUR, &newPos))
			{
				return newPos.QuadPart;
			}
		}
		return wxInvalidOffset;
	}
	wxFileOffset COMIStream::OnSysSeek(wxFileOffset pos, wxSeekMode mode)
	{
		if (m_Stream)
		{
			LARGE_INTEGER move = {};
			move.QuadPart = pos;

			ULARGE_INTEGER newPos = {};
			switch (mode)
			{
				case wxSeekMode::wxFromCurrent:
				{
					m_LastError = m_Stream->Seek(move, STREAM_SEEK_CUR, &newPos);
					break;
				}
				case wxSeekMode::wxFromStart:
				{
					m_LastError = m_Stream->Seek(move, STREAM_SEEK_SET, &newPos);
					break;
				}
				case wxSeekMode::wxFromEnd:
				{
					m_LastError = m_Stream->Seek(move, STREAM_SEEK_END, &newPos);
					break;
				}
			}

			if (m_LastError)
			{
				return newPos.QuadPart;
			}
		}
		return wxInvalidOffset;
	}
	size_t COMIStream::OnSysRead(void* buffer, size_t size)
	{
		if (m_Stream)
		{
			ULONG read = 0;
			if (m_LastError = m_Stream->Read(buffer, static_cast<ULONG>(size), &read))
			{
				return read;
			}
		}
		return 0;
	}
	size_t COMIStream::OnSysWrite(const void* buffer, size_t size)
	{
		if (m_Stream)
		{
			ULONG written = 0;
			if (m_LastError = m_Stream->Write(buffer, static_cast<ULONG>(size), &written))
			{
				return written;
			}
		}
		return 0;
	}

	bool COMIStream::IsWriteable() const
	{
		if (m_Stream)
		{
			STATSTG stat = {};
			if (m_LastError = m_Stream->Stat(&stat, STATFLAG_NONAME))
			{
				return stat.grfMode & STGM_READWRITE;
			}
		}
		return false;
	}
	bool COMIStream::IsReadable() const
	{
		if (m_Stream)
		{
			STATSTG stat = {};
			if (m_LastError = m_Stream->Stat(&stat, STATFLAG_NONAME))
			{
				return true;
			}
		}
		return false;
	}

	bool COMIStream::Flush()
	{
		if (m_Stream)
		{
			m_LastError = m_Stream->Commit(STGC_DEFAULT);
			return m_LastError.IsSuccess();
		}
		return false;
	}
	bool COMIStream::SetAllocationSize(BinarySize offset)
	{
		if (m_Stream)
		{
			ULARGE_INTEGER size = {};
			size.QuadPart = offset.GetBytes();

			m_LastError = m_Stream->SetSize(size);
			return m_LastError.IsSuccess();
		}
		return false;
	}
}
