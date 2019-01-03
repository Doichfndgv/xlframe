# XlFrame
A Python library for styling dataframes when exporting to excel.

Credit to [DeepSpace2](https://github.com/DeepSpace2) and his [StyleFrame](https://github.com/DeepSpace2/StyleFrame) package which this is based off of. 
This is just my rendition that I wrote to better fit my usage.

---

## Contents
1. [Installation](#installation)
2. [Components](#components)  
&nbsp;&nbsp;&nbsp;&nbsp;- [XlFrame](#xlframe_class)  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- [Class](#xlframe_init)  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- [Methods](#xlframe_methods)  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- [Properties](#xlframe_properties)  
&nbsp;&nbsp;&nbsp;&nbsp;- [Style](#style)  
&nbsp;&nbsp;&nbsp;&nbsp;- [utils](#utils)  
3. [Example Usage](#example-usage)    
&nbsp;&nbsp;&nbsp;&nbsp;- [Styling Data](#styling-data)  
&nbsp;&nbsp;&nbsp;&nbsp;- [Headers and Index styling](#headers-and-index-styling)  
&nbsp;&nbsp;&nbsp;&nbsp;- [Column and Row dimensions](#column-and-row-dimensions)  
&nbsp;&nbsp;&nbsp;&nbsp;- [Hyperlinks](#hyperlinks)  
&nbsp;&nbsp;&nbsp;&nbsp;- [Miscellaneous](#miscellaneous)


## Installation
```
pip install xlframe
```

## Components
<a name="xlframe_class"></a>
### XlFrame
<a name="xlframe_init"></a>
* ***Class***:
```python
# from xlframe import XlFrame

class XlFrame:

    def __init__(self, dataframe, style=None, header_style=None, index_style=None, *, number_style=None,
                 date_style=None, datetime_style=None, timedelta_style=None, use_default_formats=True):
        """
        :param dataframe: DateFrame to style.
        :type dataframe: pandas.DataFrame.
        :param style: Initial style to apply to frame. Default xlframe.Style.default_style().
        :type style: openpyxl.NamedStyle or xlframe.Style.
        :param header_style: Style for headers to use instead of style.
            Default center aligns, bolds and adds thin borders to style.
        :type header_style: openpyxl.NamedStyle or xlframe.Style.
        :param index_style: Style for index to use instead of style.
            Default bolds and adds thin borders to style. 
            Alignment and number format based on index type.
        :type index_style: openpyxl.NamedStyle or xlframe.Style.
        :param number_style: Style for numeric columns to use instead of style.
            Default right aligns style and sets number format to utils.NumberFormats.general.
        :type number_style: openpyxl.NamedStyle or xlframe.Style.
        :param date_style: Style for date columns to use instead of style.
            Default right aligns style and sets number format to utils.Options.default_date_format.
        :type date_style: openpyxl.NamedStyle or xlframe.Style.
        :param datetime_style: Style for datetime columns to use instead of style.
            Default right aligns style and sets number format to utils.Options.default_datetime_format.
        :type datetime_style: openpyxl.NamedStyle or xlframe.Style.
        :param timedelta_style: Style for timedelta columns to use instead of style.
            Default sets number format to utils.Options.default_timedelta_format.
            Tries to align left or right depending on if number format is textual or not.
        :type timedelta_style: openpyxl.NamedStyle or xlframe.Style.
        :param use_default_formats: Apply default formatting where a style is not supplied.
        :type use_default_formats: Boolean.
        """
```

Class for styling dataframes. Styling based around openpyxl's NamedStyle.
```xlframe.Style``` available as an alternative to using NamedStyle.

Pandas-like syntax for assigning styles.  
Assign to column/s:  
&nbsp;&nbsp;&nbsp;&nbsp;```my_xlframe['Column1'] = my_style```  
&nbsp;&nbsp;&nbsp;&nbsp;```my_xlframe[['Column1', 'Column2']] = my_style```  
    
Use ```.styles``` or ```.istyles``` property like you would pandas.DataFrame.loc or .iloc but to assign styles:  
&nbsp;&nbsp;&nbsp;&nbsp;```my_xlframe.styles[2:10, ['Column3', 'Column4']] = my_style```  
&nbsp;&nbsp;&nbsp;&nbsp;```my_xlframe.istyles[2:10, [2, 3]] = my_style```  
    
To edit the existing style at a set of locations pass a dictionary of changes:  
&nbsp;&nbsp;&nbsp;&nbsp;```my_xlframe.styles[2:10, ['Column3', 'Column4']] = {'bold': True, 'font_color': 'red'}```  
Accepts any kwargs ```xlframe.Style``` does.
   
```.auto_fit()``` attempts to fit column widths based on their contents.  
```.to_excel()``` to export styled dataframe. Supports .xlsx and .xlsm file formats.  

Default type specific styling adjusts alignment and number format of style arg.  
<br/>
<a name="xlframe_methods"></a>
* ***Methods***:
```python
    @staticmethod
    def ExcelWriter(path, load_existing=False, **kwargs):
        """
        :param path: Full path for workbook.
        :type path: String.
        :param load_existing: Load existing workbook into excel_writer if one exists at path.
        :type load_existing: Boolean.
        :return: pandas.ExcelWriter
        """
```

See pandas.ExcelWriter. Engine will be set to 'openpyxl'.  

---
```python
    def to_excel(self, excel_writer='output.xlsx', sheet_name='Sheet1', *, protect_sheet=False,
                 right_to_left=False, columns_to_hide=None, add_filters=False, replace_sheet=False,
                 auto_fit=None, **kwargs):
        """
        :param excel_writer: ExcelWriter or file path to export to.
        :type excel_writer: ExcelWriter or string.
        :param sheet_name: Sheet name to export to.
        :type sheet_name: string.
        :param protect_sheet: Enable protection on sheet.
        :type protect_sheet: boolean.
        :param right_to_left: Make sheet right to left oriented.
        :type right_to_left: boolean.
        :param columns_to_hide: Columns to make hidden.
        :type columns_to_hide: string, int or list-like.
        :param add_filters: Add excel filters to header row.
        :type add_filters: boolean.
        :param replace_sheet: If sheet_name already exists delete it first.
        :type replace_sheet: boolean.
        :param auto_fit: Columns to autofit. Can pass True to fit all columns.
        :type auto_fit: list-like or boolean.
        :param kwargs: Passed to pandas.DataFrame.to_excel().
        :return: pandas.ExcelWriter.
        """
```

Export to excel. Supports .xlsx/.xlsm. See pandas.DataFrame.to_excel() for more.  

---
```python
    def auto_fit(self, columns=None, scalar=None, flat=None, max_width=None, min_width=None,
                 index=True, include_header=True):
        """
        
        
        :param columns: Columns to autofit.
        :type columns: list-like.
        :param scalar: To multiply by number of characters to get width. Default utils.Options.default_autofit_scalar.
        :type scalar: float.
        :param flat: Flat amount to add to width. Default utils.Options.default_autofit_flat.
        :type flat: float.
        :param max_width: Max allowed width. Default utils.Options.default_autofit_max.
        :type max_width: float.
        :param min_width: Min allowed width. Default utils.Options.default_autofit_min.
        :type min_width: float.
        :param index: Fit index width as well.
        :type index: boolean.
        :param include_header: Also consider width of column header when fitting. For index uses index.name.
        :type include_header: boolean.
        :return: self
        """
```

Attempt to auto fit column widths. ~Max length entry in column * scalar + flat. If columns not provided fits all columns.  

---
```python
    def row_stripes(self, fill_color='D9D9D9'):
        """
        :param fill_color: Color to fill. Hex, rgb tuple or or openpyxl.styles.Color.
        :type fill_color: str, tuple or openpyxl.styles.Color
        :return: self
        """
```

Solid fill every other row with fill_color.  

---
```python
    def col_stripes(self, fill_color='D9D9D9'):
        """
        :param fill_color: Color to fill. Hex, rgb tuple or or openpyxl.styles.Color.
        :type fill_color: str, tuple or openpyxl.styles.Color
        :return: self
        """
```

Solid fill every other column with fill_color.  

---
```python
    def format_as_table(self, table_style=None, table_name=None, row_stripes=True, col_stripes=None, **kwargs):
        """
        :param table_style: Table style to use.
        :type table_style: string (Ex. 'TableStyleLight1') or openpyxl.worksheet.table.TableStyleInfo
        :param table_name: Name for table. Must be unique within workbook.
            Default will find the next available name counting up. Table1, Table2, Table3 ect.
        :type table_name: string
        :param row_stripes: Show row stripes.
        :type row_stripes: boolean
        :param col_stripes: Show col stripes.
        :type col_stripes: boolean
        :param kwargs: Passed to openpyxl.worksheet.table.Table.
        :return: self
        """
```

When exported format excel range as a table. Tables names must be unique within a workbook.  
To format as table must use headers and filters will be enabled.  

---
```python
    def add_style(self, style):
        """
        :param style: Style to add.
        :type style: xlframe.Style or openpyxl.styles.NamedStyle
        :return: Style name
        :rtype: str
        """
```

Add style to available named styles. Can be assigned just by name afterwards.     
Styles will also be automatically added when first assigned.  
<br/>
<a name="xlframe_properties"></a>
* ***Properties***:  
```python
    .styles
```

Assign styles to data. Indexes as pandas.DataFrame.loc. ```.styles[:, :] = 'MyStyle'```.  

---
```python
    .istyles
```

Assign styles to data. Indexes as pandas.DataFrame.iloc. ```.istyles[:, :] = 'MyStyle'```.  

---
```python
    .header_styles
```

Assign styles to headers. ```.header_styles = 'MyStyle'```.  
Or index as pandas.Series.  
```.header_styles[:] = 'MyStyle'``` ```.header_styles.loc[:] = 'MyStyle'``` ```.header_styles.iloc[:] = 'MyStyle'```.  

---
```python
    .index_styles
```

Assign styles to index. ```.index_styles = 'MyStyle'```.  
Or index as pandas.Series.  
```.index_styles[:] = 'MyStyle'``` ```.index_styles.loc[:] = 'MyStyle'``` ```.index_styles.iloc[:] = 'MyStyle'```.    

---
```python
    .row_heights
```

Assign row heights. ```.row_heights = 12.5```.  
Or index as pandas.Series.  
```.row_heights[:] = 12.5``` ```.row_heights.loc[:] = 12.5``` ```.row_heights.iloc[:] = 12.5```.  

---
```python
    .column_widths
```

Assign column widths. ```.column_widths = 25```.  
Or index as pandas.Series.  
```.column_widths[:] = 25``` ```.column_widths.loc[:] = 25``` ```.column_widths.iloc[:] = 25```.  

---
```python
    .header_height
```

Assign header height. ```.header_height = 12.5```.  

---
```python
    .index_width
```

Assign index width. ```.index_width = 12.5```.  

---
```python
    .hyperlinks
```

Add hyperlinks to a column. ```.hyperlinks['ColumnName'] = 'https://www.python.org/'```. See [usage examples](#hyperlinks). 

---
```python
    .loc
```

Create new XlFrame from selection of current XlFrame. ```new_xlframe = old_xlframe.loc[10:20, ['Column1', 'Column2']]```.  

---
```python
    .iloc
```

Create new XlFrame from selection of current XlFrame. ```new_xlframe = old_xlframe.iloc[10:20, 2:4]```.  

---
```python
    .builtins
```

Tuple of openpyxl's available builtin style names.  
<br/>
### Style
* ***Class***:
```python
# from xlframe import Style

class Style:

    def __init__(self, name, number_format=utils.NumberFormats.general, font_style=utils.Options.default_font_style,
                 font_size=utils.Options.default_font_size, font_color=None, bold=None, underline=None, italic=None,
                 strikethrough=None, fill_pattern='solid', fill_color=None, horizontal_alignment=None,
                 vertical_alignment=None, indent=0, wrap_text=None, shrink_to_fit=None,
                 border_style=None, border_color=None):
        """
        :param name: Style name.
        :type name: str
        :param number_format: Excel number format. See utils.NumberFormats.
        :type number_format: str
        :param font_style: Excel font style. See utils.FontStyles.
        :type font_style: str
        :param font_size: Font size.
        :type font_size: int/float
        :param font_color: Font color. See utils.Colors.
        :type font_color: Hex str, (r, g, b) tuple or openpyxl.styles.Color
        :param bold: Make font bold.
        :type bold: boolean
        :param underline: Underline font. See utils.Underline.
        :type underline: boolean
        :param italic: Italic font.
        :type italic: boolean
        :param strikethrough: strikethrough font.
        :type strikethrough: boolean
        :param fill_pattern: Excel fill pattern. See utils.FillPattern.
        :type fill_pattern: str
        :param fill_color: Fill color. See utils.Colors.
        :type fill_color: Hex str, (r, g, b) tuple or openpyxl.styles.Color
        :param horizontal_alignment: Excel horizontal alignment. See utils.Alignments.Horizontal.
        :type horizontal_alignment: str
        :param vertical_alignment: Excel vertical alignment. See utils.Alignments.Vertical.
        :type vertical_alignment: str
        :param indent: Cell indent.
        :type indent: int
        :param wrap_text: Enable wrap text.
        :type wrap_text: boolean
        :param shrink_to_fit: Enable shrink to fit.
        :type shrink_to_fit: boolean
        :param border_style: Border style. See utils.BorderStyles.
        :type border_style: str
        :param border_color: Border color. See utils.Colors.
        :type border_color: Hex str, (r, g, b) tuple or openpyxl.styles.Color
        """
```

Optional constructor for styles.  
Style.to_named_style() or Style.named_style to get equivalent openpyxl.styles.NamedStyle.  
Style.from_named_style(named_style) to convert openpyxl.styles.NamedStyle to equivalent Style.  
<br/>
* ***Methods***:
```python
    def to_named_style(self):
        """
        :return: openpyxl.styles.NamedStyle 
        """
```

Convert to equivalent openpyxl.styles.NamedStyle.  

---
```python
    @classmethod
    def from_named_style(cls, style):
        """
        :param style: openpyxl.styles.NamedStyle 
        :return: Style
        """
```

Convert openpyxl.styles.NamedStyle to equivalent Style.  

---
* ***Properties***: 

```python
    .name
    .number_format
    .font_style
    .font_size
    .font_color
    .bold
    .underline
    .italic
    .strikethrough
    .fill_pattern
    .fill_color
    .horizontal_alignment
    .vertical_alignment
    .indent
    .wrap_text
    .shrink_to_fit
    .border_style
    .border_color
    .locked
    .hidden
```

### utils
* ***utils***: 
```python
from xlframe import utils

utils.NumberFormats
utils.BorderStyles
utils.FillPattern
utils.Underline
utils.Alignments.Vertical
utils.Alignments.Horizontal
utils.Colors
utils.FontStyles
utils.Options
```

Various constants and options for default settings.  

```python
from xlframe.utils import Options

Options.default_date_format
Options.default_time_format
Options.default_datetime_format
Options.default_timedelta_format

Options.default_font_size
Options.default_font_style

Options.default_autofit_scalar
Options.default_autofit_flat
Options.default_autofit_min
Options.default_autofit_max

Options.default_column_width
Options.default_row_height
```

## Example Usage

### Styling Data
```python
import pandas as pd
from openpyxl.worksheet.table import TableStyleInfo

from xlframe import XlFrame, Style, utils

df = pd.DataFrame({
    'Dates': [pd.datetime(year=2018, month=1, day=1) + pd.Timedelta(days=i) for i in range(10)],
    'Datetimes': [pd.datetime(year=2018, month=1, day=1) + pd.Timedelta(minutes=i * 100) for i in range(10)],
    'Strings': ['Abc', 'Def', 'Ghi', 'Jkl', 'Mno', 'Pqr', 'Stu', 'Vwx', 'Yz0', '123'],
    'Numbers': [i * 12345 for i in range(1, 11)],
})

# Simple export with default styling.
xf = XlFrame(df)
"""
Underlying dataframe.
       Dates           Datetimes Strings  Numbers
0 2018-01-01 2018-01-01 00:00:00     Abc    12345
1 2018-01-02 2018-01-01 01:40:00     Def    24690
2 2018-01-03 2018-01-01 03:20:00     Ghi    37035
3 2018-01-04 2018-01-01 05:00:00     Jkl    49380
4 2018-01-05 2018-01-01 06:40:00     Mno    61725
5 2018-01-06 2018-01-01 08:20:00     Pqr    74070
6 2018-01-07 2018-01-01 10:00:00     Stu    86415
7 2018-01-08 2018-01-01 11:40:00     Vwx    98760
8 2018-01-09 2018-01-01 13:20:00     Yz0   111105
9 2018-01-10 2018-01-01 15:00:00     123   123450

Styles
          Dates         Datetimes  Strings         Numbers
0  Default Date  Default Datetime  Default  Default Number
1  Default Date  Default Datetime  Default  Default Number
2  Default Date  Default Datetime  Default  Default Number
3  Default Date  Default Datetime  Default  Default Number
4  Default Date  Default Datetime  Default  Default Number
5  Default Date  Default Datetime  Default  Default Number
6  Default Date  Default Datetime  Default  Default Number
7  Default Date  Default Datetime  Default  Default Number
8  Default Date  Default Datetime  Default  Default Number
9  Default Date  Default Datetime  Default  Default Number
"""
# Export with auto_fit and filters. Filters go on top row.
xf.to_excel('test.xlsx', sheet_name='test', auto_fit=True, add_filters=True, index=False)
# Only auto_fit certain columns
xf.to_excel('test.xlsx', sheet_name='test', auto_fit=xf.columns[:2], add_filters=True)

# Different ways to assign styles. Pandas-like syntax.
xf = XlFrame(df)
# Create example style
style = Style(
    name='Example',
    number_format='dd MMMM yyyy HH:mm',
    border_style=utils.BorderStyles.thin,
    border_color=utils.Colors.blue
)

# Assign style to entire column.
xf['Dates'] = style
# or columns.
xf[['Dates', 'Datetimes']] = 'Example'  # First assignment will register name. Can assign by just name after.

"""
Styles
     Dates Datetimes  Strings         Numbers
0  Example   Example  Default  Default Number
1  Example   Example  Default  Default Number
2  Example   Example  Default  Default Number
3  Example   Example  Default  Default Number
4  Example   Example  Default  Default Number
5  Example   Example  Default  Default Number
6  Example   Example  Default  Default Number
7  Example   Example  Default  Default Number
8  Example   Example  Default  Default Number
9  Example   Example  Default  Default Number
"""

# .styles property indexes as dataframe.loc but is for assigning styles.
xf.styles[xf['Numbers'] > 50000, ['Strings', 'Numbers']] = 'Neutral'  # Built-in style
# __getitem__ operations like (xf['SomeNumbers'] > 50000) will access
# the dataframe so you can do indexing like the above
# or source dataframe accessible at xf.dataframe.

# .istyles property indexes as dataframe.iloc but is for assigning styles.
xf.istyles[5:, [0, 1]] = 'Bad'

"""
Styles
     Dates Datetimes  Strings         Numbers
0  Example   Example  Default  Default Number
1  Example   Example  Default  Default Number
2  Example   Example  Default  Default Number
3  Example   Example  Default  Default Number
4  Example   Example  Neutral         Neutral
5      Bad       Bad  Neutral         Neutral
6      Bad       Bad  Neutral         Neutral
7      Bad       Bad  Neutral         Neutral
8      Bad       Bad  Neutral         Neutral
9      Bad       Bad  Neutral         Neutral
"""

style2 = Style(
    name='Example2',
    font_size=16,
    font_color=utils.Colors.turquoise
)
# Convert xlframe.Style to openpyxl.styles.NamedStyle
named_style2 = style2.named_style
# Register ahead of time with XlFrame
xf.add_style(named_style2)
# Assign by name
xf.styles[8:, 'Numbers'] = 'Example2'
# Assigning the object also still works
xf.styles[7, 'Numbers'] = named_style2
# Or the original xlframe.Style object
xf.styles[1, 'Numbers'] = style2

"""
Styles
     Dates Datetimes  Strings         Numbers
0  Example   Example  Default  Default Number
1  Example   Example  Default        Example2
2  Example   Example  Default  Default Number
3  Example   Example  Default  Default Number
4  Example   Example  Neutral         Neutral
5      Bad       Bad  Neutral         Neutral
6      Bad       Bad  Neutral         Neutral
7      Bad       Bad  Neutral        Example2
8      Bad       Bad  Neutral        Example2
9      Bad       Bad  Neutral        Example2
"""

# Pass dictionary of modifications for existing styles at cell locations.
# Accepts any properties xlframe.Style does.
# Resulting new styles will be given a new unique (numbered) name and registered.
xf.styles[:5, ['Datetimes', 'Strings']] = {
    'fill_pattern': utils.FillPattern.solid, 'fill_color': utils.Colors.light_orange
}
"""
Styles
     Dates   Datetimes     Strings         Numbers
0  Example  Example[1]  Default[1]  Default Number
1  Example  Example[1]  Default[1]        Example2
2  Example  Example[1]  Default[1]  Default Number
3  Example  Example[1]  Default[1]  Default Number
4  Example  Example[1]  Neutral[1]         Neutral
5      Bad      Bad[1]  Neutral[1]         Neutral
6      Bad         Bad     Neutral         Neutral
7      Bad         Bad     Neutral        Example2
8      Bad         Bad     Neutral        Example2
9      Bad         Bad     Neutral        Example2
"""

# To format the resulting excel range as a table
xf.format_as_table()
# Default just formats it as a table. No additional styling.
# To add table styles pass name of a table style
xf.format_as_table('TableStyleMedium3', row_stripes=False)
# Or create and pass openpyxl.worksheet.table.TableStyleInfo
table_style = TableStyleInfo('TableStyleLight1', showColumnStripes=True)
xf.format_as_table(table_style)
```   

### Headers and Index styling
```python
import pandas as pd

from xlframe import XlFrame, Style

df = pd.DataFrame({
    'Dates': [pd.datetime(year=2018, month=1, day=1) + pd.Timedelta(days=i) for i in range(10)],
    'Datetimes': [pd.datetime(year=2018, month=1, day=1) + pd.Timedelta(minutes=i * 100) for i in range(10)],
    'Strings': ['Abc', 'Def', 'Ghi', 'Jkl', 'Mno', 'Pqr', 'Stu', 'Vwx', 'Yz0', '123'],
    'Numbers': [i * 12345 for i in range(1, 11)],
})
    
xf = XlFrame(df)

# Header Styles. Defaults:
"""
Dates        Default Header
Datetimes    Default Header
Strings      Default Header
Numbers      Default Header
Name: HeaderStyles, dtype: object
"""
# Index Styles. Defaults:
"""
0    Default Number Index
1    Default Number Index
2    Default Number Index
3    Default Number Index
4    Default Number Index
5    Default Number Index
6    Default Number Index
7    Default Number Index
8    Default Number Index
9    Default Number Index
Name: IndexStyles, dtype: object
"""

# Supports similar kinds of assignments as Data.
style = Style(name='Example')
named_style = Style(name='Example2').named_style

# Assign to all headers
xf.header_styles = style
# Or index as pandas.Series.
xf.header_styles[0:2] = 'Good'
xf.header_styles['Numbers'] = 'Bad'
xf.header_styles.loc[['Numbers', 'Datetimes']] = 'Neutral'
xf.header_styles.iloc[-1] = named_style

"""
Dates            Good
Datetimes     Neutral
Strings       Example
Numbers      Example2
Name: HeaderStyles, dtype: object
"""

xf.index_styles = 'Neutral'
xf.index_styles[5:] = 'Good'
xf.index_styles[7] = 'Bad'
xf.index_styles.loc[1:3] = style
xf.index_styles.iloc[[3, 4, 5]] = named_style

"""
0     Neutral
1     Example
2     Example
3    Example2
4    Example2
5    Example2
6        Good
7         Bad
8        Good
9        Good
Name: IndexStyles, dtype: object
"""
```    


### Column and Row dimensions

```python
import pandas as pd

from xlframe import XlFrame

df = pd.DataFrame({
    'Dates': [pd.datetime(year=2018, month=1, day=1) + pd.Timedelta(days=i) for i in range(10)],
    'Datetimes': [pd.datetime(year=2018, month=1, day=1) + pd.Timedelta(minutes=i * 100) for i in range(10)],
    'Strings': ['Abc', 'Def', 'Ghi', 'Jkl', 'Mno', 'Pqr', 'Stu', 'Vwx', 'Yz0', '123'],
    'Numbers': [i * 12345 for i in range(1, 11)],
})
    
xf = XlFrame(df)

# xf.column_widths. Default width = utils.Options.default_column_width
"""
Dates        8.43
Datetimes    8.43
Strings      8.43
Numbers      8.43
Name: ColumnWidths, dtype: float64
"""
# xf.row_heights. Default height = utils.Options.default_row_height
"""
0    15.0
1    15.0
2    15.0
3    15.0
4    15.0
5    15.0
6    15.0
7    15.0
8    15.0
9    15.0
Name: RowHeights, dtype: float64
"""

# Assign to all
xf.column_widths = 13.5
# Or index as pandas.Series.
xf.row_heights[:7] = 22

"""
Dates          13.5
Datetimes      13.5
Strings        13.5
SomeNumbers    13.5
Name: ColumnWidths, dtype: float64

0    22.0
1    22.0
2    22.0
3    22.0
4    22.0
5    22.0
6    22.0
7    15.0
8    15.0
9    15.0
Name: RowHeights, dtype: float64
"""

# Header height and index width stored separately as individual numbers.
xf.header_height = 20
xf.index_width = 16
   

```

### Hyperlinks
```python
import pandas as pd
from openpyxl.worksheet.hyperlink import Hyperlink

from xlframe import XlFrame

df = pd.DataFrame({
    'Dates': [pd.datetime(year=2018, month=1, day=1) + pd.Timedelta(days=i) for i in range(10)],
    'Datetimes': [pd.datetime(year=2018, month=1, day=1) + pd.Timedelta(minutes=i * 100) for i in range(10)],
    'Strings': ['Abc', 'Def', 'Ghi', 'Jkl', 'Mno', 'Pqr', 'Stu', 'Vwx', 'Yz0', '123'],
    'Numbers': [i * 12345 for i in range(1, 11)],
})
    
xf = XlFrame(df)

# To add hyperlinks to a column use the .hyperlinks property
# Can assign as a simple string
xf.hyperlinks['Strings'] = 'https://exampleurl.com/'
# Or as an openpyxl.worksheet.hyperlink.Hyperlink
xf.hyperlinks['Numbers'] = Hyperlink(
    ref='',  # ref will be assigned for you when to_excel() is called
    target='https://exampleurl.com/',
    tooltip='abcdefghijklmnopqrstuvwxyz',
)
xf.hyperlinks['Numbers'] = xf['Numbers'].apply(
    lambda x: Hyperlink(
        ref='',  # ref will be assigned for you when to_excel() is called
        target='https://exampleurl.com/?x={}'.format(x),
        tooltip='My Number is {}'.format(x)
    )
)

# Hyperlinks won't automatically be styled as hyperlinks.
# To add hyperlink styling
xf['Numbers'] = 'Hyperlink'

                    
```

### Miscellaneous
```python
import pandas as pd

from xlframe import XlFrame

df = pd.DataFrame({
    'Dates': [pd.datetime(year=2018, month=1, day=1) + pd.Timedelta(days=i) for i in range(10)],
    'Datetimes': [pd.datetime(year=2018, month=1, day=1) + pd.Timedelta(minutes=i * 100) for i in range(10)],
    'Strings': ['Abc', 'Def', 'Ghi', 'Jkl', 'Mno', 'Pqr', 'Stu', 'Vwx', 'Yz0', '123'],
    'Numbers': [i * 12345 for i in range(1, 11)],
})

# Disable default styling
xf = XlFrame(df, style='Normal', use_default_formats=False)
"""
Styles
    Dates Datetimes Strings Numbers
0  Normal    Normal  Normal  Normal
1  Normal    Normal  Normal  Normal
2  Normal    Normal  Normal  Normal
3  Normal    Normal  Normal  Normal
4  Normal    Normal  Normal  Normal
5  Normal    Normal  Normal  Normal
6  Normal    Normal  Normal  Normal
7  Normal    Normal  Normal  Normal
8  Normal    Normal  Normal  Normal
9  Normal    Normal  Normal  Normal
"""

# Provide type specific initial styles
xf = XlFrame(
    df, style='Normal', header_style='Pandas', index_style='Pandas', 
    number_style='Good', date_style='Bad', datetime_style='Neutral'
)
"""
Styles
  Dates Datetimes Strings Numbers
0   Bad   Neutral  Normal    Good
1   Bad   Neutral  Normal    Good
2   Bad   Neutral  Normal    Good
3   Bad   Neutral  Normal    Good
4   Bad   Neutral  Normal    Good
5   Bad   Neutral  Normal    Good
6   Bad   Neutral  Normal    Good
7   Bad   Neutral  Normal    Good
8   Bad   Neutral  Normal    Good
9   Bad   Neutral  Normal    Good
"""

# Slice an XlFrame with .loc or .iloc
xf2 = xf.loc[2:7, ['Dates', 'Datetimes', 'Strings']]
"""
Styles
  Dates Datetimes Strings
2   Bad   Neutral  Normal
3   Bad   Neutral  Normal
4   Bad   Neutral  Normal
5   Bad   Neutral  Normal
6   Bad   Neutral  Normal
7   Bad   Neutral  Normal
"""

xf3 = xf.iloc[5:, [0, 1]]
"""
Styles
  Dates Datetimes
5   Bad   Neutral
6   Bad   Neutral
7   Bad   Neutral
8   Bad   Neutral
9   Bad   Neutral
"""                    
```
