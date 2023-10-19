# **Style a Pandas Dataframe Using openpyxl**
## style-pandas-dataframe

Class made to style a DataFrame using a template excel file, coping all the attributes, like font, text color, function, data validation, format.

### Parameters
> file_template_path: template path
> save_path: final save path
> file_path: dataframe saved as excel file
> adictional_rules_to_apply: another conditional formatattion 
```
{"<SHEET_NAME>": [
        {"<COLUMN_NAME>": Rule}
    ]
}
```
> datas_validation_to_apply: add a data validation (add a select input in cells)
 ```
 self.datas_validation_to_apply = {
    "<SHEETNAME>":{
        "<CELLS_APPLY>": "<FORMULA>"
    }
}
 ```
> show_gridlines: Remove grid border if False