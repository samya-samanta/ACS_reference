# ACS Reference Formatter

Formats numbered bibliography entries from a Word (`.docx`) document to the requested American Chemical Society (ACS) presentation, **without changing article titles or author names**.

- Journal titles abbreviated and italicized using a curated abbreviation dictionary.
- Publication years **bold** and volume numbers *italic*.
- Page ranges use en dashes (–).
- The default demonstration modifies **only the first 10 references**.
- Unknown or irregular journal references are left unchanged and flagged for review; **do not assume all journal names are covered**.

## Run locally

```bash
pip install python-docx
python format_acs.py REFERENCES.docx --output ACS_demo.docx --limit 10
```

For all entries after reviewing the pilot, use `--limit 0`.

## Run on GitHub

1. Upload `REFERENCES.docx` to the repository root.
2. Open **Actions → Format ACS references → Run workflow**.
3. Enter `10` for the sample size and download the `acs-formatted-references` artifact after completion.

## Editorial rules

The dictionary distinguishes *Angew. Chem.* (German edition) from *Angew. Chem., Int. Ed.*; volume 129 (2017) in the example is the **German edition**. Consult the [CAS Source Index (CASSI)](https://cassi.cas.org/) for definitive abbreviations when extending the dictionary. Historical journal titles may need manual verification.

The formatter is a pilot automation, not a guarantee that all entries are valid; DOI, author spellings, page ranges and metadata are preserved from the input, not fact-checked.
