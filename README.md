# VBA_GAS_UDFs

## Summary

General utility functions for Excel and Google Sheets written in in VBA and GAS.
These are generally call as __user-defined__ or __custom__ functions.
The Excel VBA functions are generally stored in _Personal.xlsb_ and are thereby available to all Excel files. Google Sheets does not have such a mechanism for sharing such functions so they have to be copied into the GAS editor of any Sheets documents where they are to be used.


## Contents guide

### VBA

There are descriptions of each of the functions in the file _GeneralFunctions.bas_. 

#### Function: _URL_

__Summary__:Extracts the URL from cell hyperlinks, returns an empty string if the input cell is not a hyperlink.

__Usage__: Provides a convenient way to get the URL from cells with hyperlinks which is slightly painful from to do from the Excel interface.

#### Function: _DATE_AS_ISO8601_STRING_

__Summary__: Returns the date of a cell as an ISO8601-formatted string or an empty string if the input cell is not of type date.

__Usage__: Excel (and Sheets for that matter) requires caution when working with dates. Cell values can be automatically converted to dates and this can lead to data corruption. Also, dates are presented differently in different locales; for example, is the date _2/1/2001_ the second of January 2001 or the first of February 2001? It depends on the locale; in the UK the date would be interpreted as the former but as the latter in the US. This function can be used to convert a date to a recognised international standard: [ISO8601](https://en.wikipedia.org/wiki/ISO_8601). Furthermore, it stored the date as a string which serves to protect it from further Excel re-interpretation and stores it as a value that is easily interpreted by other storage mechanisms such as relational databases or when serialised as a JSON string.

## GAS

