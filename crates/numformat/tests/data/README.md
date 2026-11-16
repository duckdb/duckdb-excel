# Regression corpora

Each line of a `.tsv` file is one case: `format<TAB>value<TAB>expected`. There are no comment or header lines.
The expectations were recorded from the formatter this crate replaced, built for x86-64 Linux, and the crate is
required to reproduce them byte for byte. They record behaviour, not intent: odd results are expected results.

- `golden.tsv`: holds on every platform.
- `golden_glibc.tsv`: cases whose digits depend on how the C library rounds `log10` and `pow` (values just below a
  power of ten, or of very large or small magnitude). They are checked on x86-64 Linux with glibc only.

## Fields

`format` and `expected` are byte strings written as printable ASCII:

| Written | Byte |
| --- | --- |
| `\\` | backslash |
| `\t`, `\n`, `\r` | tab, line feed, carriage return |
| `\xHH` | the byte with hex value `HH`; used for bytes below 0x20, 0x7F and all bytes from 0x80 |
| anything else | itself |

Formats are valid UTF-8 once decoded. Results need not be: a format can cut a multi-byte character apart, and
`NumberFormat::format_bytes` returns such results unchanged.

`expected` is the single escape `\E` where the formatter gives up (`FormatError`; `format_number` then returns
an empty string).

`value` is a decimal number as Rust parses it (`1234.5`, `1e-7`, `-0`, `inf`, `nan`), or `0x` followed by the
16 hex digits of the IEEE 754 bit pattern for values a decimal cannot express, such as NaNs with a sign or payload.
