CommcareTranslationChecker
==========================

https://github.com/dimagi/CommcareTranslationChecker

A command-line tool to check multiple columns of a [Bulk Translation File](https://confluence.dimagi.com/display/commcarepublic/Form+Bulk+Translation) against each other, to ensure that output value tags are being used consistently between columns.

Installation
--------------------------

0a\. Install Python and `pip`. This tool is tested with Python 2.7, and 3.6.

1\. Install CommCare Translation Checker via `pip`

```
$ pip install CommcareTranslationChecker
```


Basic Command-line Usage
------------------------

The basic usage of the command-line tool is with a saved Excel file. This can either be configured for [Form Translation](https://confluence.dimagi.com/display/commcarepublic/Form+Bulk+Translation) or [Application Translation](https://confluence.dimagi.com/display/commcarepublic/Bulk+Application+Translations)

```
$ CommcareTranslationChecker  <relative or absolute path to translation file>
```

By default, this will read the specified file, and check all columns whose names start with "default_" against the left-most "default_" column. If any discrepancies are found between the list of "output value" tags in any of the columns, a file will be generated in the folder "commcareTranslationChecker_output." If no such folder exists relative to the current path, it will be created. This file will be an exact copy of the data in the input file, with an additional column "mismatchFlag" appended to each sheet. This column will be flagged "Y" in all rows for which a disprepancy was detected, and "N" otherwise. In addition, all cells whose "output value" tags differ from the left-most column's will be red-filled, for easy visual reference.

If the translation file contains a sheet called Modules_and_forms, with a column called sheet_names, the tool will check that each value in this column corresponds to the name of one of the sheets in the workbook. If not, the corresponding cell in the sheet_names column of the output file will be highlighted red.

After the file has been created, a summary will be printed outlining how many rows were found to have discrepancies per sheet.


Use via import
------------------------
```
>>> import openpyxl
>>> from CommcareTranslationChecker import validate_workbook
>>> messages = []
>>> wb = openpyxl.load_workbook("examples/TranslationCheckerTest_BulkAppTranslation.xlsx")
>>> validate_workbook(wb, messages)
Traceback (most recent call last):
  File "<stdin>", line 1, in <module>
  File "CommcareTranslationChecker/CommcareTranslationChecker.py", line 527, in validate_workbook
    raise FatalError("Some fatal error message.")
CommcareTranslationChecker.exceptions.FatalError: Some fatal error message.
>>> messages
['There were issues with the following worksheets:', u'moduleX_formY is missing from the workbook.']
```

Advanced Command-line Usage
---------------------------
In addition to the basic usage outlined, there are a number of optional parameters that will provide a more customized experience.

```
$ CommcareTranslationChecker    --columns <comma-separated list of column names to check> \
                                --base-column <name of column to be compared against, if different from left-most> \
                                --output-folder <relative or absolute path to folder in which to save output file> \
                                --ignore-order \
                                --verbose \
                                --no-output-file \
                                --configuration-sheet <name of sheet containing meta information about other sheets in the workbook> \
                                --configuration-sheet-column <name of column in configuration-sheet containing expected sheet titles> \
                                --output-mismatch-types \
                                --format-check \
                                --format-check-characters <sequence of characters whose counts will be compared by format check> \
                                --format-check-characters-add <sequence of characters to add to the current format check character list> \

                                
```

The five options that do not include an input parameter are described below:
* **--ignore-order** If passed, the order in which output value tags appear will not be considered when comparing cells against each other. This is useful if the order of the output value tags is different between columns because of differences in word orders between the languages involved.
* **--verbose** If passed, output will be printed to the screen pointing out which rows of the file have issues.
* **--no-output-file** If passed, no output file will be created.
* **--output-mismatch-types** If passed, will include further information about the mismatch in the output. If an output file is generated, this information will be appended as an additional column on each sheet for each language column that contains an error. If the **--verbose** flag is passed, this information will be added to each line of output.
* **--format-check** If passed, will add an additional check to compare the count of any special characters between columns. The default character list is ~`!@#$%^&*()_-+={[}]|\\:;\"'<,>.?/

See `CommcareTranslationChecker --help` for the full list of options.



Release process
---------------

1\. Create a tag for the release

```
$ git tag -a "X.YY.0" -m "Release X.YY.0"
$ git push --tags
```

2\. Create the source distribution

Ensure that the archive has the correct version number (matching the tag name).
```
$ python setup.py sdist
```

3\. Create the wheel
```
$ python setup.py bdist_wheel --universal
```

4\. Upload to pypi

```
$ pip install twine
$ twine upload dist/CommcareTranslationChecker-X.YY.0*
```

5\. Verify upload

https://pypi.python.org/pypi/CommcareTranslationChecker

6\. Create a release on github

https://github.com/dimagi/CommcareTranslationChecker/releases