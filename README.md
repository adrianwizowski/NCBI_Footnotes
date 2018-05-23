# NCBI Footnotes
Program inspired by genetics students from University of Wrocław.
Creates to find urls to NCBI PubMed database and change them to a scientific footnotes in docx or txt files.
Works only for urls from https://www.ncbi.nlm.nih.gov/pubmed/

## How does it work?
Module parses Word or text file searching for urls to NCBI PubMed database.
Articles ids from parsed urls are send as a query to PubMed using biopython module.
From results recieveid from PubMed scientific footnotes are build.

Footnotes will be created using pattern:
List of authors: AuthorLastName AuthorInitial., comma separated, publication date in brackets, full title,
ISO standardized journal title, The volume number of the journal in which the article was published,
issue - part or supplement of the journal in which the article was published in brackets (if available),
International Standard Serial Number (if available).

For example:
```
https://www.ncbi.nlm.nih.gov/pubmed/15698585 <- number is the article id.
```
will be changed to:
```
King RE., Kent KD., Bomser JA. (2005) Resveratrol reduces oxidation and proliferation of human retinal pigment
epithelial cells via extracellular signal-regulated kinase inhibition. Chem. Biol. Interact. 151(2): 143-9
```

## Usage.
### Footnotes.py
Run module using parse_txt() function for text files with NCBI urls or
run module using parse_docx() function for Word files.

### parse_txt()
Run function, choose input file and name of output file.
Function will save both url and footnotes in output file.

### parse_docx()
Run function, choose input file and name of output file. For safety do not overwrite input file.
Function will change paragraphs containing urls to scientific footnotes in output file.
Leading number of a footnote from file (under 1000) will stay in place.

Example:
```
95 https://www.ncbi.nlm.nih.gov/pubmed/26434591
```
will be changed to:
```
95 Gan L., Xiu R., Ren P., Yue M., Su H., Guo G., Xiao D., Yu J., Jiang H., Liu H., Hu G., Qing G. (2016)
Metabolic targeting of oncogene MYC by selective activation of the proton-coupled monocarboxylate family of transporters.
Oncogene 35(23): 3037-48
```
No hyperlinks or scoring allowed. Url in hyperlinks or scoring points will remain untouched.

### Dependencies:
* [biopython](https://github.com/biopython)
* [python-docx](https://python-docx.readthedocs.io/en/latest/)
* [tkinter](https://wiki.python.org/moin/TkInter)
