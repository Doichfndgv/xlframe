import locale as _locale

from openpyxl.styles import DEFAULT_FONT as _DEFAULT_FONT
from openpyxl.styles.builtins import styles as _styles
from openpyxl.styles.colors import COLOR_INDEX


class NumberFormats:
    text = '@'
    general = 'General'
    general_integer = '0'
    general_float = '0.00'
    percent = '0.0%'
    thousands_comma_sep = '#,##0'

    date = 'MM/DD/YYYY' if _locale.getdefaultlocale()[0] == 'en_US' else 'DD/MM/YYYY'
    date_long = 'mmm dd, yyyy'

    time_24_hours = 'HH:MM'
    time_24_hours_with_seconds = 'HH:MM:SS'
    time_12_hours = 'h:MM AM/PM'
    time_12_hours_with_seconds = 'h:MM:SS AM/PM'

    date_time = '{} {}'.format(date, time_24_hours)
    date_time_12_hours = '{} {}'.format(date, time_12_hours)
    date_time_with_seconds = '{} {}'.format(date, time_24_hours_with_seconds)
    date_time_12_hours_with_seconds = '{} {}'.format(date, time_12_hours_with_seconds)

    timedelta_fractional_days = general_float
    timedelta_fractional_days_with_text = timedelta_fractional_days + ' "days"'
    timedelta_24plus_hours = '[h]:mm'
    timedelta_24plus_hours_with_seconds = '[h]:mm:ss'
    timedelta_days_hours = 'd "days" h:mm'
    timedelta_days_hours_with_seconds = 'd "days" h:mm:ss'


class Colors:
    # https://github.com/ClosedXML/ClosedXML/wiki/Excel-Indexed-Colors
    black = COLOR_INDEX[0]
    white = COLOR_INDEX[1]
    red = COLOR_INDEX[2]
    bright_green = COLOR_INDEX[3]
    blue = COLOR_INDEX[4]
    yellow = COLOR_INDEX[5]
    pink = COLOR_INDEX[6]
    turquoise = COLOR_INDEX[7]
    # black = COLOR_INDEX[8]
    # white = COLOR_INDEX[9]
    # red = COLOR_INDEX[10]
    # bright_green = COLOR_INDEX[11]
    # blue = COLOR_INDEX[12]
    # yellow = COLOR_INDEX[13]
    # pink = COLOR_INDEX[14]
    # turquoise = COLOR_INDEX[15]
    dark_red = COLOR_INDEX[16]
    green = COLOR_INDEX[17]
    dark_blue = COLOR_INDEX[18]
    dark_yellow = COLOR_INDEX[19]
    violet = COLOR_INDEX[20]
    teal = COLOR_INDEX[21]
    grey_25 = COLOR_INDEX[22]
    grey_50 = COLOR_INDEX[23]
    periwinkle = COLOR_INDEX[24]
    plum = COLOR_INDEX[25]
    ivory = COLOR_INDEX[26]
    light_turquoise = COLOR_INDEX[27]
    dark_purple = COLOR_INDEX[28]
    coral = COLOR_INDEX[29]
    ocean_blue = COLOR_INDEX[30]
    ice_blue = COLOR_INDEX[31]
    # dark_blue = COLOR_INDEX[32]
    # pink = COLOR_INDEX[33]
    # yellow = COLOR_INDEX[34]
    # turquoise = COLOR_INDEX[35]
    # violet = COLOR_INDEX[36]
    # dark_red = COLOR_INDEX[37]
    # teal = COLOR_INDEX[38]
    # blue = COLOR_INDEX[39]
    sky_blue = COLOR_INDEX[40]
    # light_turquoise = COLOR_INDEX[41]
    light_green = COLOR_INDEX[42]
    light_yellow = COLOR_INDEX[43]
    pale_blue = COLOR_INDEX[44]
    rose = COLOR_INDEX[45]
    lavender = COLOR_INDEX[46]
    tan = COLOR_INDEX[47]
    light_blue = COLOR_INDEX[48]
    aqua = COLOR_INDEX[49]
    lime = COLOR_INDEX[50]
    gold = COLOR_INDEX[51]
    light_orange = COLOR_INDEX[52]
    orange = COLOR_INDEX[53]
    blue_grey = COLOR_INDEX[54]
    grey_40 = COLOR_INDEX[55]
    dark_teal = COLOR_INDEX[56]
    sea_green = COLOR_INDEX[57]
    dark_green = COLOR_INDEX[58]
    olive_green = COLOR_INDEX[59]
    brown = COLOR_INDEX[60]
    # plum = COLOR_INDEX[61]
    indigo = COLOR_INDEX[62]
    grey_80 = COLOR_INDEX[63]


class FontStyles:
    arial = 'Arial'
    baskerville = 'Baskerville Old Face'
    bodoni = 'Bodoni MT'
    calibri = 'Calibri'
    consolas = 'Consolas'
    courier_new = 'Courier New'
    garamond = 'Garamond'
    gill_sans = 'Gill Sans MT'
    leelawadee = 'Leeawadee'
    lucida_console = 'Lucida Console'
    rockwell = 'Rockwell'
    segoe_ui = 'Segoe UI'
    tahoma = 'Tahoma'
    times_new_roman = 'Times New Roman'
    trebuchet_ms = 'Trebuchet MS'
    tw_cen_mt = 'Tw Cen MT'
    verdana = 'Verdana'


class BorderStyles:
    dash_dot = 'dashDot'
    dash_dot_dot = 'dashDotDot'
    dashed = 'dashed'
    dotted = 'dotted'
    double = 'double'
    hair = 'hair'
    medium = 'medium'
    medium_dash_dot = 'mediumDashDot'
    medium_dash_dot_dot = 'mediumDashDotDot'
    medium_dashed = 'mediumDashed'
    slant_dash_dot = 'slantDashDot'
    thick = 'thick'
    thin = 'thin'


class Alignments:

    class Horizontal:
        general = 'general'
        left = 'left'
        center = 'center'
        right = 'right'
        fill = 'fill'
        justify = 'justify'
        center_continuous = 'centerContinuous'
        distributed = 'distributed'

    class Vertical:
        top = 'top'
        center = 'center'
        bottom = 'bottom'
        justify = 'justify'
        distributed = 'distributed'


class Underline:
    single = 'single'
    double = 'double'


class FillPattern:
    solid = 'solid'
    dark_down = 'darkDown'
    dark_gray = 'darkGray'
    dark_grid = 'darkGrid'
    dark_horizontal = 'darkHorizontal'
    dark_trellis = 'darkTrellis'
    dark_up = 'darkUp'
    dark_vertical = 'darkVertical'
    gray0625 = 'gray0625'
    gray125 = 'gray125'
    light_down = 'lightDown'
    light_gray = 'lightGray'
    light_grid = 'lightGrid'
    light_horizontal = 'lightHorizontal'
    light_trellis = 'lightTrellis'
    light_up = 'lightUp'
    light_vertical = 'lightVertical'
    medium_gray = 'mediumGray'


class StyleEdits:
    bad = {
        'font_color': _styles['Bad'].font.color,
        'fill_color': _styles['Bad'].fill.fgColor,
        'fill_pattern': _styles['Bad'].fill.patternType
    }
    good = {
        'font_color': _styles['Good'].font.color,
        'fill_color': _styles['Good'].fill.fgColor,
        'fill_pattern': _styles['Good'].fill.patternType
    }
    normal = {
        'font_color': _styles['Normal'].font.color,
        'fill_color': _styles['Normal'].fill.fgColor,
        'fill_pattern': _styles['Normal'].fill.patternType
    }
    neutral = {
        'font_color': _styles['Neutral'].font.color,
        'fill_color': _styles['Neutral'].fill.fgColor,
        'fill_pattern': _styles['Neutral'].fill.patternType
    }
    highlight = {
        'font_color': '00FF0000',
        'fill_color': '00FFFF00',
        'fill_pattern': FillPattern.solid
    }


class Options:
    default_date_format = NumberFormats.date
    default_time_format = NumberFormats.time_24_hours_with_seconds
    default_datetime_format = NumberFormats.date_time_with_seconds
    default_timedelta_format = NumberFormats.timedelta_fractional_days

    default_font_size = 11
    default_font_style = 'Calibri'
    _DEFAULT_FONT.sz = default_font_size

    default_autofit_scalar = 1.25
    default_autofit_flat = 1.5
    default_autofit_min = 6.86
    default_autofit_max = 150

    default_column_width = 8.43
    default_row_height = 15


# openpyxl uses 12 as its font size for built-ins but excel seems to default to 11. Set these to 11.
for _, _style in _styles.items():
    _style.font.sz = Options.default_font_size


if __name__ == '__main__':
    pass
