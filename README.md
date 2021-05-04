# csv_word_merge
Merge CSV fields into a MS Word template

**Note:** I threw this together in a couple of hours to help a friend, so it is not as complete as my other projects.
* Error checking could be improved.
* Using `concurrent.futures` would also be nice to have to process files faster.

## Description
This program allows you to merge rows in a CSV file into a MS Word document.  A PDF file is also saved.

* The first row of the CSV file should contain a header.
* Example: `Email,First,Last`

* The MS Word document should then contain, case-sensative *macros* with underscores:
* `_Email_`
* `_First_`
* `_Last_`

Created files are saved via the `-C` switch (note the capital `C`).  If your CSV file has a field called `Email`, then you could use `-C Email`.  There would then be two newly created files for `user@example.com`:
* `user@example.com.docx`
* `user@example.com.pdf`

## Requirements

* tested with `Python 3.9`
* pip install python-docx
* pip install docx2pdf

## Usage
```
usage: csv_word_merge.py [-h] --csv CSV --col COL --dest DEST [--version]
                         wordfile

Merge CSV fields into a MS Word template

positional arguments:
  wordfile              MS Word file with macros

optional arguments:
  -h, --help            show this help message and exit
  --csv CSV, -c CSV     csv file containing macros
  --col COL, -C COL     column name for output PDF
  --dest DEST, -d DEST  destination folder
  --version, -v         display version and then exit

```

## Example

* csv file: `clients.csv`:

| ID | First
|----|-----|
| 12 | Bubba |


* col: use a column named `ID` to name the output files
* dest: save both `docx` and `pdf` files to this directory, in this case a `surveys` folder
* word document: `template.docx`, which contains a `_First_` macro

```
python3 csv_word_merge.py --csv clients.csv --col ID --dest surveys template.docx
```

* In the `surveys` directoy, you should have 2 files, with `_First_` substituted out for `Bubba`
* `12.docx`
* `12.pdf`

## LICENSE
* [MIT License](LICENSE)
