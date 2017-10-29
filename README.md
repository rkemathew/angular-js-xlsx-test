There are 4 .xlsx files included in this repo for testing.

DataTemplate -> Should end with success
DataTemplateColumnMap -> Should end with unmapped errors displayed on screen
DataTemplateMissingFields -> Should end with list of empty fields
DataTemplateSalaryReviewDate -> Should end with no salary reviewed date provided message.


There are 3 json files included in this repo.

column_names.json has all of the accepted template column names. (could feasibly do without this by using the keys on the column_to_mapping.json in its place)
column_to_mapping.json has a mapping of the column names to the mapping field
error_check_info.json has fields determining the error checks to be done.  This should be built from the survey responses