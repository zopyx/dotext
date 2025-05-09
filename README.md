Document File Text Extractor
=============================

[![Build Status](https://travis-ci.org/anvie/dotext.svg?branch=master)](https://travis-ci.org/anvie/dotext)
[![Build status](https://ci.appveyor.com/api/projects/status/rghm59ie4ax9655t?svg=true)](https://ci.appveyor.com/project/anvie/dotext)
[![Crates.io](https://img.shields.io/crates/v/dotext.svg)](https://crates.io/crates/dotext)

Simple Rust library to extract readable text from specific document format like Word Document (docx).
Currently only support several format, other format coming soon.

Supported Document
-------------------------

- [x] Microsoft Word (docx)
- [x] Microsoft Excel (xlsx)
- [x] Microsoft Power Point (pptx)
- [x] OpenOffice Writer (odt)
- [x] OpenOffice Spreadsheet (ods)
- [x] OpenDocument Presentation (odp)
- [ ] PDF

Usage
------

### As a Library

```rust
let mut file = Docx::open("samples/sample.docx").unwrap();
let mut isi = String::new();
let _ = file.read_to_string(&mut isi);
println!("CONTENT:");
println!("----------BEGIN----------");
println!("{}", isi);
println!("----------EOF----------");
```

### Command-line Tool

The package includes a command-line tool called `extract-docx` for extracting text from DOCX files:

```bash
# Write output to a file
extract-docx --from sample.docx --to output.txt

# Write output to stdout
extract-docx --from sample.docx --stdout

# Using short options
extract-docx -f sample.docx -t output.txt
extract-docx -f sample.docx -s
```

Test
-----

```bash
$ cargo test
```

or run example:

```bash
$ cargo run --example readdocx data/sample.docx
```

[] Robin Sy.
