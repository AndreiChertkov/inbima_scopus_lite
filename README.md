# inbima_scopus_lite


## Description

**IN**telligent system for **BI**bliographic data **MA**nagement from Scopus database (lite version). This software package allows to export the bibliographic list of the author's publications in a structured, neat form, broken down by quartiles.


## How to use

1. Download this [repository](https://github.com/AndreiChertkov/inbima_scopus_lite).
2. Open downloaded folder `inbima_scopus_lite` in terminal (console) and install dependencies by the command (`python 3.7+` is required) `pip install -r requirements.txt`.
3. Open author profile in Scopus (e.g., [this one](https://www.scopus.com/authid/detail.uri?authorId=8529104000)) and push `Export all` (`Экспортировать все` in russian) button, then enable checkboxes for all types of information and select `csv` format for export. Rename the downloaded file as `[surname].csv` and move it into the `data` folder in the repo.
4. Run the main script by the command `python inbima_scopus_lite.py` with optional argument, which relates to lower bound for the year of publication (e.g., `python inbima_scopus_lite.py 2017`). If the argument is not provided, then all publications will be exported.
5. The folder `result` will contain all automatically generated reports in the docx-format: lists of publications for each of the authors separately (`result_[surname].docx`), as well as a consolidated list (`result.docx`).
    > Each file contains publications from the first and second quartiles in separate sections, as well as other publications within a single numbering. If necessary, section headings can be removed, which will allow you to get a complete numbered list of publications.

> Some biblio entries contain incorrect journal names. Name mapping can be specified in the `data/journals_map.py` file. Note that a similar problem exists on the Scopus website (the link to the journal for some papers is not working).

> Some biblio entries contain invalid papers. Skipable titles can be specified in the `data/papers_del.py` file.


## How to update the journals database (optional)

Parsed information about journals indexed in the Scopus database is contained in a neat structured form in the file `data/journals.xlsx` in the root of the project (last update: **May 15, 2022**). Several times a year (specifically, 3 times), the information in the Scopus database is updated. To take into account new data, you should perform the following steps:

1. Go to [Scopus](https://www.scopus.com/sources) and log in to the system (you may need access on behalf of the organization to be able to export data).
2. Click the `Download Scopus Source List` (`Скачать список источников Scopus` in russian) button above the journals table, select the `Download source titles and metrics` (`Скачать названия источника и показатели` in russian) option, save and unpack the archive, getting a file like `CiteScore-2011-2020-new-methodology-October-2021`. Save the base creation date somewhere (for example, `October 2021`).
3. Open the downloaded excel file, for the target sheet (sheet with a name like `CiteScore 2020`) select the `Move or copy` option (`Переместить или скопировать` in russian) and specify copying to a new workbook. Name the target sheet in the new file as `data`. Similarly, copy the `ASJC codes` sheet into the same new document and name it `codes`. Save the new file as `scopus.xlsx` in the standard `xlsx` format, then move this file to the `data` folder of the project. The original excel file can then be deleted.
4. Run the script to parse the journals by the command `python journals_parse.py`. After successful execution of the script, make sure that the `data/journals.xlsx` file has appeared (or updated) and contains a valid number of records.

> Note that the number of saved records in the downloaded excel file (`scopus.xlsx`) will be greater than the number provided in Scopus web-site, since the journals in the table are duplicated if there are several areas of knowledge for the journal with the corresponding quartiles. The number of records in the resulting excel file (`data/journals.xlsx`) will be even less, since incorrectly specified sources (for example, if ISSN and EISSN are not set at the same time) are not included in the list. Also, we do not include journals with duplicate titles in the list. See warnings printed to the console when the script is running. For problematic journals included in the final list (for example, those with a duplicate title), an appropriate comment will be given in the "Error" column.


## Author

- [Andrei Chertkov](https://github.com/AndreiChertkov) (a.chertkov@skoltech.ru).
