#include "xlsx/zip_file.hpp"
#include "xlsx/xml_util.hpp"

#include "duckdb/common/file_system.hpp"

#include "minizip-ng/mz.h"
#include "minizip-ng/mz_os.h"
#include "minizip-ng/mz_strm.h"
#include "minizip-ng/mz_zip.h"
#include "minizip-ng/mz_zip_rw.h"

namespace duckdb {

//-------------------------------------------------------------------------
// Minizip File Stream
//-------------------------------------------------------------------------

// NOLINTBEGIN
int32_t mz_stream_duckdb_open(void *stream, const char *path, int32_t mode);
int32_t mz_stream_duckdb_is_open(void *stream);
int32_t mz_stream_duckdb_read(void *stream, void *buf, int32_t size);
int32_t mz_stream_duckdb_write(void *stream, const void *buf, int32_t size);
int64_t mz_stream_duckdb_tell(void *stream);
int32_t mz_stream_duckdb_seek(void *stream, int64_t offset, int32_t origin);
int32_t mz_stream_duckdb_close(void *stream);
int32_t mz_stream_duckdb_error(void *stream);
void *mz_stream_duckdb_create(void);
void mz_stream_duckdb_delete(void **stream);

static mz_stream_vtbl mz_duckdb_file_stream_vtable = {
    mz_stream_duckdb_open,
    mz_stream_duckdb_is_open,
    mz_stream_duckdb_read,
    mz_stream_duckdb_write,
    mz_stream_duckdb_tell,
    mz_stream_duckdb_seek,
    mz_stream_duckdb_close,
    mz_stream_duckdb_error,
    mz_stream_duckdb_create,
    mz_stream_duckdb_delete,
    nullptr,
    nullptr,
};

struct mz_stream_duckdb {
	mz_stream base;
	FileSystem *fs;
	FileHandle *handle;
	string last_error;
};

int32_t mz_stream_duckdb_open(void *stream, const char *path, int32_t mode) {
	auto &self = *reinterpret_cast<mz_stream_duckdb *>(stream);

	if (self.handle) {
		self.handle->Close();
		self.handle->~FileHandle();
		self.handle = nullptr;
		self.last_error.clear();
	}

	FileOpenFlags flags = 0;
	if (mode & MZ_OPEN_MODE_READ) {
		flags |= FileFlags::FILE_FLAGS_READ;
	}
	if (mode & MZ_OPEN_MODE_WRITE) {
		flags |= FileFlags::FILE_FLAGS_WRITE;
	}
	if (mode & MZ_OPEN_MODE_APPEND) {
		flags |= FileFlags::FILE_FLAGS_APPEND;
	}
	if (mode & MZ_OPEN_MODE_CREATE) {
		flags |= FileFlags::FILE_FLAGS_FILE_CREATE;
	}
	if (mode & MZ_OPEN_MODE_EXISTING) {
		flags |= FileFlags::FILE_FLAGS_READ;
	}
	try {
		auto file = self.fs->OpenFile(path, flags);
		if (!file) {
			return MZ_OPEN_ERROR;
		}
		self.handle = file.release();
	} catch (Exception &ex) {
		ErrorData err(ex);
		self.last_error = err.RawMessage();
		return MZ_OPEN_ERROR;
	}
	return MZ_OK;
}

int32_t mz_stream_duckdb_is_open(void *stream) {
	auto &self = *reinterpret_cast<mz_stream_duckdb *>(stream);
	if (!self.handle) {
		return MZ_OPEN_ERROR;
	}
	return MZ_OK;
}

int32_t mz_stream_duckdb_read(void *stream, void *buf, int32_t size) {
	auto &self = *reinterpret_cast<mz_stream_duckdb *>(stream);
	return self.handle->Read(buf, size);
}

int32_t mz_stream_duckdb_write(void *stream, const void *buf, int32_t size) {
	auto &self = *reinterpret_cast<mz_stream_duckdb *>(stream);
	return self.handle->Write(const_cast<void *>(buf), size);
}

int64_t mz_stream_duckdb_tell(void *stream) {
	auto &self = *reinterpret_cast<mz_stream_duckdb *>(stream);
	return self.handle->SeekPosition();
}

int32_t mz_stream_duckdb_seek(void *stream, int64_t offset, int32_t origin) {
	auto &self = *reinterpret_cast<mz_stream_duckdb *>(stream);
	switch (origin) {
	case MZ_SEEK_SET:
		self.handle->Seek(offset);
		break;
	case MZ_SEEK_CUR:
		self.handle->Seek(self.handle->SeekPosition() + offset);
		break;
	case MZ_SEEK_END:
		// is the offset negative when seeking from the end?
		self.handle->Seek(self.handle->SeekPosition() + offset);
		break;
	default:
		return MZ_SEEK_ERROR;
	}
	return MZ_OK;
}

int32_t mz_stream_duckdb_close(void *stream) {
	auto &self = *reinterpret_cast<mz_stream_duckdb *>(stream);
	self.handle->Close();
	self.handle->~FileHandle();
	self.handle = nullptr;
	self.last_error.clear();
	return MZ_OK;
}

int32_t mz_stream_duckdb_error(void *stream) {
	// TODO: DuckDB's fs tend to throw on errors, so this is not needed
	auto &self = *reinterpret_cast<mz_stream_duckdb *>(stream);
	return mz_stream_error(&self.base);
}

void *mz_stream_duckdb_create(void) {
	auto stream = new mz_stream_duckdb();
	stream->base.vtbl = &mz_duckdb_file_stream_vtable;
	return stream;
}

void mz_stream_duckdb_delete(void **stream) {
	if (!stream) {
		return;
	}
	auto &self = *reinterpret_cast<mz_stream_duckdb *>(*stream);
	if (self.handle) {
		self.handle->Close();
		self.handle->~FileHandle();
		self.handle = nullptr;
		self.last_error.clear();
	}
	delete reinterpret_cast<mz_stream_duckdb *>(*stream);
	*stream = nullptr;
}
// NOLINTEND

//-------------------------------------------------------------------------
// Zip File Writer
//-------------------------------------------------------------------------

ZipFileWriter::ZipFileWriter(ClientContext &context, const string &file_name) {
	handle = mz_zip_writer_create();
	stream = mz_stream_duckdb_create();
	is_entry_open = false;

	auto &fs = FileSystem::GetFileSystem(context);

	auto &duckdb_stream = *static_cast<mz_stream_duckdb *>(stream);
	duckdb_stream.fs = &fs;
	duckdb_stream.handle = nullptr;

	// Set compression level
	mz_zip_writer_set_compress_level(handle, MZ_COMPRESS_LEVEL_NORMAL);
	mz_zip_writer_set_compress_method(handle, MZ_COMPRESS_METHOD_DEFLATE);

	if (mz_stream_open(stream, file_name.c_str(), MZ_OPEN_MODE_CREATE | MZ_OPEN_MODE_WRITE) != MZ_OK) {
		if (duckdb_stream.last_error.empty()) {
			throw IOException("Failed to open file for writing");
		} else {
			throw IOException(duckdb_stream.last_error);
		}
	}

	if (mz_zip_writer_open(handle, stream, 0) != MZ_OK) {
		if (duckdb_stream.last_error.empty()) {
			throw IOException("Failed to open zip for writing");
		} else {
			throw IOException(duckdb_stream.last_error);
		}
	}
}

ZipFileWriter::~ZipFileWriter() {
	if (handle) {
		if (mz_zip_writer_is_open(handle)) {
			mz_zip_writer_close(handle);
		}
		mz_zip_writer_delete(&handle);
	}
	if (stream) {
		if (mz_stream_is_open(stream)) {
			mz_stream_close(stream);
		}
		mz_stream_delete(&stream);
	}
}

void ZipFileWriter::AddDirectory(const string &dir_path) {
	// Must end with a slash
	D_ASSERT(dir_path[dir_path.size() - 1] == '/');
	mz_zip_file file_info = {};
	file_info.filename = dir_path.c_str();
	file_info.compression_method = MZ_COMPRESS_METHOD_STORE;
	file_info.zip64 = MZ_ZIP64_AUTO; // only use 64-bit when necessary

	mz_zip_writer_add_buffer(handle, nullptr, 0, &file_info);
}

void ZipFileWriter::BeginFile(const string &file_path) {
	if (is_entry_open) {
		throw IOException("ZipWriter: Cannot open a new entry before closing the previous one");
	}
	mz_zip_file file_info = {0};
	file_info.filename = file_path.c_str();
	file_info.compression_method = MZ_COMPRESS_METHOD_DEFLATE;
	file_info.zip64 = MZ_ZIP64_AUTO; // only use 64-bit when necessary

	file_info.uncompressed_size = 0;
	file_info.compressed_size = 0;

	if (mz_zip_writer_entry_open(handle, &file_info) != MZ_OK) {
		throw IOException("Failed to open entry for writing");
	}
	is_entry_open = true;
}

idx_t ZipFileWriter::Write(const char *str) {
	return Write(str, strlen(str));
}

idx_t ZipFileWriter::Write(const string &str) {
	return Write(str.c_str(), str.size());
}

idx_t ZipFileWriter::Write(const char *buffer, idx_t write_size) {
	auto bytes_written = mz_zip_writer_entry_write(handle, buffer, write_size);
	if (bytes_written < 0) {
		throw IOException("Failed to write entry");
	}
	return bytes_written;
}

void ZipFileWriter::EndFile() {
	if (!is_entry_open) {
		throw IOException("ZipWriter: Cannot close an entry that is not open");
	}
	if (mz_zip_writer_entry_close(handle) != MZ_OK) {
		throw IOException("Failed to close entry");
	}
	is_entry_open = false;
}

void ZipFileWriter::Finalize() {
	if (mz_zip_writer_is_open(handle)) {
		mz_zip_writer_close(handle);
	}
	if (mz_stream_is_open(stream)) {
		mz_stream_close(stream);
	}
}

//-------------------------------------------------------------------------
// Zip File Reader
//-------------------------------------------------------------------------

ZipFileReader::ZipFileReader(ClientContext &context, const string &file_name) {
	handle = mz_zip_reader_create();
	stream = mz_stream_duckdb_create();
	is_entry_open = false;
	entry_pos = 0;
	entry_len = 0;

	auto &fs = FileSystem::GetFileSystem(context);

	auto &duckdb_stream = *static_cast<mz_stream_duckdb *>(stream);
	duckdb_stream.fs = &fs;
	duckdb_stream.handle = nullptr;

	if (mz_stream_open(stream, file_name.c_str(), MZ_OPEN_MODE_READ) != MZ_OK) {
		if (duckdb_stream.last_error.empty()) {
			throw IOException("Failed to open file for reading");
		} else {
			throw IOException(duckdb_stream.last_error);
		}
	}

	if (mz_zip_reader_open(handle, stream) != MZ_OK) {
		if (duckdb_stream.last_error.empty()) {
			throw IOException("Failed to open zip for reading");
		} else {
			throw IOException(duckdb_stream.last_error);
		}
	}
}

bool ZipFileReader::TryOpenEntry(const string &file_name) {
	if (mz_zip_reader_locate_entry(handle, file_name.c_str(), 0) != MZ_OK) {
		return false;
	}

	if (mz_zip_reader_entry_open(handle) != MZ_OK) {
		return false;
	}

	mz_zip_file *file_info = nullptr;
	if (mz_zip_reader_entry_get_info(handle, &file_info) != MZ_OK) {
		return false;
	}

	const auto len = file_info->uncompressed_size;

	is_entry_open = true;
	entry_pos = 0;
	entry_len = len;

	return true;
}

void ZipFileReader::CloseEntry() {
	if (!is_entry_open) {
		throw IOException("ZipReader: Cannot close an entry that is not open");
	}
	const auto close_result = mz_zip_reader_entry_close(handle);
	if (close_result != MZ_OK) {
		const auto is_early_exit = close_result == MZ_CRC_ERROR && entry_pos < entry_len;

		// Allow CRC error if we close before reading the entire entry
		if (!is_early_exit) {
			throw IOException("Failed to close entry");
		}
	}
	is_entry_open = false;
	entry_pos = 0;
}

idx_t ZipFileReader::Read(char *buffer, const idx_t read_size) {
	const auto bytes_read = mz_zip_reader_entry_read(handle, buffer, static_cast<int32_t>(read_size));
	if (bytes_read < 0) {
		throw IOException("Failed to read entry");
	}
	entry_pos += bytes_read;
	return bytes_read;
}

idx_t ZipFileReader::GetEntryPos() const {
	return entry_pos;
}

idx_t ZipFileReader::GetEntryLen() const {
	return entry_len;
}

bool ZipFileReader::IsDone() const {
	return entry_pos >= entry_len;
}

ZipFileReader::~ZipFileReader() {
	if (handle) {
		if (mz_zip_reader_is_open(handle)) {
			mz_zip_reader_close(handle);
		}
		mz_zip_reader_delete(&handle);
	}
	if (stream) {
		if (mz_stream_is_open(stream)) {
			mz_stream_close(stream);
		}
		mz_stream_delete(&stream);
	}
}

} // namespace duckdb
