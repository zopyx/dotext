/**
 * Copyright 2017 Robin Syihab. All rights reserved.
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this software
 * and associated documentation files (the "Software"), to deal in the Software without restriction,
 * including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense,
 * and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so,
 * subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all copies
 * or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
 * INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
 * PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
 * FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE,
 * ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
 * IN THE SOFTWARE.
 *
 */
extern crate clap;
extern crate dotext;

use clap::{App, Arg};
use dotext::Docx;
use dotext::MsDoc;
use std::fs::File;
use std::io::{self, Read, Write};

fn main() -> io::Result<()> {
    let matches = App::new("extract-docx")
        .version("0.1.0")
        .about("Extracts text from a DOCX file")
        .arg(
            Arg::with_name("from")
                .long("from")
                .short("f")
                .takes_value(true)
                .required(true)
                .help("Source DOCX file path"),
        )
        .arg(
            Arg::with_name("to")
                .long("to")
                .short("t")
                .takes_value(true)
                .help("Output file path"),
        )
        .arg(
            Arg::with_name("stdout")
                .long("stdout")
                .short("s")
                .help("Write output to stdout"),
        )
        .get_matches();

    let source_path = matches.value_of("from").unwrap();

    // Verify the source file exists and is a docx file
    if !source_path.ends_with(".docx") {
        return Err(io::Error::new(
            io::ErrorKind::InvalidInput,
            "Source file must be a .docx file",
        ));
    }

    // Open the docx file
    let mut docx = match Docx::open(source_path) {
        Ok(doc) => doc,
        Err(e) => {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                format!("Failed to open DOCX file: {}", e),
            ));
        }
    };

    // Read the content
    let mut content = String::new();
    docx.read_to_string(&mut content)?;

    // Verify we have either --stdout or --to
    if !matches.is_present("stdout") && matches.value_of("to").is_none() {
        return Err(io::Error::new(
            io::ErrorKind::InvalidInput,
            "Either --to or --stdout must be specified",
        ));
    }

    // Check if we should output to stdout
    if matches.is_present("stdout") {
        print!("{}", content);
        return Ok(());
    }

    // Otherwise, write to the output file
    if let Some(output_path) = matches.value_of("to") {
        let mut file = match File::create(output_path) {
            Ok(f) => f,
            Err(e) => {
                return Err(io::Error::new(
                    io::ErrorKind::Other,
                    format!("Failed to create output file: {}", e),
                ));
            }
        };

        match file.write_all(content.as_bytes()) {
            Ok(_) => println!("Successfully extracted text to {}", output_path),
            Err(e) => {
                return Err(io::Error::new(
                    io::ErrorKind::Other,
                    format!("Failed to write to output file: {}", e),
                ));
            }
        }
    }

    Ok(())
}