"""
pdf_parser_ASAIL_Public.py
==========================
Extract text from a local PDF of a judicial decision and output a
structured plain-text file compatible with the corpus format expected
by analyze_cases_ASAIL_Public.py.

USAGE
-----
  python pdf_parser_ASAIL_Public.py <input.pdf> [output.txt]

  If output.txt is omitted, the file is written alongside the PDF with
  the same name but a .txt extension.

OUTPUT FORMAT
-------------
The output file uses the same section delimiter format as the full
corpus:

    -----
    HEADNOTES
    -----
    [headnote placeholder — fill in manually or leave blank]
    -----
    CORE JUDGMENT
    -----
    [full judgment text extracted from PDF]
    -----
    FOOTNOTES
    -----
    [footnotes extracted from PDF, if any]

FOOTNOTE HANDLING
-----------------
Superscript footnote reference numbers (e.g. "opinion¹⁵ of the court")
are stripped using two patterns:
  1. Digit(s) immediately after a letter, before whitespace/punctuation:
       r'(?<=[a-zA-Z])(\d{1,2})(?=[\s.,;:!\?\)\]\"])'
  2. Year+footnote run-ons (e.g. "[2015]15" from PDF encoding):
       r'(\b(?:18|19|20)\d{2})(\d{1,2})(?=[\s.,;:\?\!])'

REQUIREMENTS
------------
  pip install pdfplumber
"""

import re
import sys
import argparse
from pathlib import Path

try:
    import pdfplumber
except ImportError:
    sys.exit(
        "ERROR: pdfplumber is not installed.\n"
        "Install with:  pip install pdfplumber"
    )

# ---------------------------------------------------------------------------
# Superscript footnote reference stripping
# ---------------------------------------------------------------------------

# Pattern 1: digit(s) attached to end of a word, before space/punctuation
# e.g. "court15 held" → "court held"
_SUP_AFTER_WORD = re.compile(
    r'(?<=[a-zA-Z])(\d{1,2})(?=[\s.,;:!\?\)\]\"\'\\])'
)

# Pattern 2: year followed immediately by a footnote digit run-on
# e.g. "[2015]15 CLR" → "[2015] CLR"
_SUP_YEAR_RUNON = re.compile(
    r'(\b(?:18|19|20)\d{2})(\d{1,2})(?=[\s.,;:\?\!])'
)


def strip_superscript_refs(text):
    """Remove superscript footnote reference numbers from extracted text."""
    text = _SUP_YEAR_RUNON.sub(r'\1', text)
    text = _SUP_AFTER_WORD.sub('', text)
    return text


# ---------------------------------------------------------------------------
# PDF extraction
# ---------------------------------------------------------------------------

def extract_pdf_text(pdf_path):
    """Extract all text from a PDF using pdfplumber.

    Returns a list of (page_number, page_text) tuples.
    """
    pages = []
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages, 1):
            txt = page.extract_text() or ''
            pages.append((i, txt))
    return pages


def clean_page_text(text):
    """Basic cleanup of a single page's extracted text."""
    # Normalise line endings
    text = text.replace('\r\n', '\n').replace('\r', '\n')
    # Remove page headers/footers that are pure page numbers
    text = re.sub(r'^\s*\d+\s*$', '', text, flags=re.MULTILINE)
    # Collapse runs of blank lines to one
    text = re.sub(r'\n{3,}', '\n\n', text)
    return text.strip()


def pdf_to_corpus_file(pdf_path, output_path=None):
    """Convert a PDF to a corpus-format .txt file.

    Parameters
    ----------
    pdf_path  : str or Path  — input PDF
    output_path : str or Path or None — output .txt (default: same dir, .txt extension)
    """
    pdf_path = Path(pdf_path)
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF not found: {pdf_path}")

    if output_path is None:
        output_path = pdf_path.with_suffix('.txt')
    output_path = Path(output_path)

    print(f"Extracting: {pdf_path.name}", flush=True)
    pages = extract_pdf_text(pdf_path)
    print(f"  {len(pages)} pages extracted", flush=True)

    # Combine all pages
    full_text = '\n\n'.join(clean_page_text(pt) for _, pt in pages if pt.strip())
    full_text = strip_superscript_refs(full_text)

    # Try to detect if footnotes appear at the end
    # Common pattern: "FOOTNOTES\n1.  ..." or a separator line before numbered notes
    fn_split = re.search(
        r'\n[-─]{5,}\s*\n\s*(?:FOOTNOTES?|Notes?|Endnotes?)\s*\n|'
        r'\n\s*(?:FOOTNOTES?|Notes?|Endnotes?)\s*\n[-─]{5,}',
        full_text, re.IGNORECASE
    )

    if fn_split:
        core_part = full_text[:fn_split.start()].strip()
        fn_part   = full_text[fn_split.end():].strip()
    else:
        core_part = full_text.strip()
        fn_part   = ''

    # Build output in corpus format
    lines = [
        '-----',
        'HEADNOTES',
        '-----',
        '[Headnotes not extracted from PDF — fill in manually if required]',
        '-----',
        'CORE JUDGMENT',
        '-----',
        core_part,
    ]
    if fn_part:
        lines += [
            '-----',
            'FOOTNOTES',
            '-----',
            fn_part,
        ]

    output_path.write_text('\n'.join(lines), encoding='utf-8')
    print(f"  Written → {output_path}", flush=True)
    return output_path


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description='Extract text from a local judicial PDF into corpus format.'
    )
    parser.add_argument('input',  help='Input PDF file')
    parser.add_argument('output', nargs='?', default=None,
                        help='Output .txt file (default: same name as PDF with .txt extension)')
    args = parser.parse_args()
    pdf_to_corpus_file(args.input, args.output)


if __name__ == '__main__':
    main()
