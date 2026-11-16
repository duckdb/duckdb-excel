.PHONY: clean clean_all

PROJ_DIR := $(dir $(abspath $(lastword $(MAKEFILE_LIST))))

EXTENSION_NAME=excel

# Mark the extension as C_STRUCT_UNSTABLE: DuckDB 2.x requires the _init_c_api_v2 entrypoint.
# Binaries load only in the exact TARGET_DUCKDB_VERSION, not in newer DuckDB releases.
USE_UNSTABLE_C_API=1
CAPI_VERSION=2

# Target DuckDB version
TARGET_DUCKDB_VERSION=v2.0.0

all: configure debug

# Include makefiles from DuckDB
include extension-ci-tools/makefiles/c_api_extensions/base.Makefile
include extension-ci-tools/makefiles/c_api_extensions/rust.Makefile

# Build workspace libraries rather than the CI tools' Wasm example target.
IS_EXAMPLE=
ifneq ($(DUCKDB_WASM_PLATFORM),)
	TARGET_INFO=--target $(TARGET) --package excel_wasm
else
	TARGET_INFO+=--package excel
endif

# The runner writes the files of __TEST_DIR__ into its test directory, so keep that out of test/sql.
TEST_SCRATCH_DIR=duckdb_unittest_tempdir
TEST_LIST=$(TEST_SCRATCH_DIR)/tests.txt
TEST_RUNNER_BASE=rm -rf $(TEST_SCRATCH_DIR) && mkdir -p $(TEST_SCRATCH_DIR) && find test/sql -name '*.test' | sort > $(TEST_LIST) && \
	$(TEST_RUNNER) --test-dir $(TEST_SCRATCH_DIR) --file-list $(TEST_LIST) $(EXTRA_EXTENSIONS_PARAM)

configure: venv platform extension_version

debug: build_extension_library_debug build_extension_with_metadata_debug
release: build_extension_library_release build_extension_with_metadata_release

test: test_debug
test_debug: test_extension_debug
test_release: test_extension_release

clean: clean_build clean_rust
clean_all: clean_configure clean
