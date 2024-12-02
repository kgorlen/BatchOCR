<!--
Markdown Guide: https://www.markdownguide.org/basic-syntax/
-->
<!--
Disable markdownlint errors:
fenced-code-language MD040
no-inline-html MD033
-->
<!-- markdownlint-disable MD040 MD033-->

# batchocr

**batchocr** - Batch convert PDF files to searchable PDF files using ABBYY FineReader Sprint

# SYNOPSIS

**batchocr** [**-h** | **--help**] [**-d** | **--debug** | **--no-debug**]
[**--dpi** *DPI*] [**-f** | **--force** | **--no-force**]   [**--pages**
*PAGES*] [**--words** *WORDS*] [**-q** | **--quiet**   | **--no-quiet**]
  [**--image** *IMAGE*] [**--text** *TEXT*] [**-a** | **--analyze** |
  **--no-analyze** | **--commit** | **--no-commit** | **--rollback** |
  **--no-rollback**] [**-v** | **--version** | **--no-version**]
  [*FILES* ...]

# DESCRIPTION

**batchocr** analyzes the first pages of the listed PDF files and writes
searchable PDFs to `*_OCR_.pdf` files if any analyzed page is not searchable.

The **--analyze** option classifies the first **--pages** files of a PDF dcoument as
*Blank*, *Unsearchable* or *Searchable* without converting the document.  A page
is classified as *Blank* if there no text blocks, images, or drawings on the
page.  Otherwise, the percentage of page area covered by text and images is
calculated and the page is classified as *Unsearchable* if the text area is less
than or equal to **--text** or if the image area is greater than **--image**.

Text blocks are not included in the covered area unless they contain at least
**--words** of five or more characters.

**NOTE:** The **--analyze** option is not reliable.  The **--pages**,
**--text**, **--image**, and **--words** options can be adjusted to improve
accuracy.

The **--commit** option deletes the unsearchable PDF files and renames the
`*_OCR_.pdf` files to replace them.

The **--rollback** option deletes the `*_OCR_.pdf` files.

**batchocr** also writes a log file named **batchocr.log** to the conventional
OS-dependent log directory,
`C:\Users\`*`Username`*`\AppData\Local\BatchOCR\Logs` on Windows.

# OPTIONS

**-h, --help**
: Print a help message and exit.

**-d, --debug, --no-debug**
: Log debugging information; default --no-debug.

**--dpi** *DPI*
: Resolution of PDF to TIFF conversion; default 300 dpi.

**-f, --force, --no-force**
: Convert pdf file even if searchable; default False.

**--pages** *PAGES*
: Maximum number of pages to analyze for searchability; default 10.

**--words** *WORDS*
: Minimum number of words in text box for inclusion in text area; default 3.

**-q, --quiet, --no-quiet**
: Do not print INFO, WARNING, ERROR, or CRITICAL messages to `stderr`; default --no-quiet.

**--image** *IMAGE*
: Page unsearchable if image area exceeds this percent; default 100.

**--text** *TEXT*
: Page unsearchable unless text area exceeds this percent; default 5.

**-a, --analyze, --no-analyze**
: Analyze PDF pages and report without conversion; default False.

**--commit, --no-commit**
: Replace original .pdf files with searchable `*_OCR_.pdf` files; default False.

**--rollback, --no-rollback**
: Delete `*_OCR_.pdf` files; default False.

**-v, --version, --no-version**
: Display the version number and exit.

# ARGUMENTS

*FILES*
: list of files and/or directories with globbing, environment variable
    substitution, `~` expansion,  and optional `**` recursion; files are read
    from stdin if "-".

# INSTALLATION

## PREREQUISITES

[Install python 3.12 or later version](https://www.python.org/downloads/).

Install [pipx](https://pipx.pypa.io/stable/):

```
pip install pipx
```

Install [ABBYY FineReader Sprint 9.0](https://archive.org/details/abbyy-fine-reader-9).

## INSTALL **batchocr** FROM `.whl` package

<pre>
<code>pipx install <i>path</i>\batchocr-<i>version</i>-py3-none-any.whl</code>
</pre>

For example:

<pre>
<code>pipx install <i>path</i>\batchocr-0.1.5-py3-none-any.whl</code>
</pre>

## INSTALL **batchocr** FROM `.tar.gz` package

Alternatively, install **batchocr** from a `.tar.gz` package file:

<pre>
<code>pipx install <i>path</i>\batchocr-<i>version</i>.tar.gz</code>
</pre>

For example:

<pre>
<code>pipx install <i>path</i>\batchocr-0.1.5-.tar.gz</code>
</pre>

# SEE ALSO

* [ABBYY -- Saving the results via the command line](https://help.abbyy.com/en-us/finereader/15/user_guide/commandline_save)<br>
* [ABBYY FineReader Sprint 9.0 Download](https://archive.org/details/abbyy-fine-reader-9)<br>
* [5 Python libraries to convert PDF to Images](https://levelup.gitconnected.com/4-python-libraries-to-convert-pdf-to-images-7a09eba83a09)<br>
* [How to check if PDF is scanned image or contains text](https://stackoverflow.com/questions/55704218/how-to-check-if-pdf-is-scanned-image-or-contains-text)<br>
* [How can I distinguish a digitally-created PDF from a searchable PDF?](https://stackoverflow.com/questions/63494812/how-can-i-distinguish-a-digitally-created-pdf-from-a-searchable-pdf)<br>
* [Analysing English -- Basic analysis](https://www.petercollingridge.co.uk/blog/language/analysing-english/basic-analysis)<br>

# AUTHOR

Keith Gorlen<br>
<kgorlen@gmail.com>

# COPYRIGHT

The MIT License (MIT)

Copyright (c) 2024 Keith Gorlen

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.
