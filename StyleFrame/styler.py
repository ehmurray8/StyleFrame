# coding:utf-8
import json
from pprint import pformat

from colour import Color
from openpyxl.comments import Comment
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import PatternFill, NamedStyle, Color as OpenPyColor, Border, Side, Font, Alignment, Protection

from . import utils


class Styler(object):
    """
    Creates openpyxl Style to be applied
    """

    cache = {}
    style_count = 0

    def __init__(self, bg_color=None, bold=False, font=utils.fonts.arial, font_size=12, font_color=None,
                 number_format=utils.number_formats.general, protection=False, underline=None,
                 border_type=utils.borders.thin, horizontal_alignment=utils.horizontal_alignments.center,
                 vertical_alignment=utils.vertical_alignments.center, wrap_text=True, shrink_to_fit=True,
                 fill_pattern_type=utils.fill_pattern_types.solid, indent=0, comment_author=None, comment_text=None):

        def get_color_from_string(color_str, default_color=None):
            if color_str and color_str.startswith('#'):
                color_str = color_str[1:]
            if not utils.is_hex_color_string(hex_string=color_str):
                color_str = utils.colors.get(color_str, default_color)
            return color_str

        self.bold = bold
        self.font = font
        self.font_size = font_size
        self.number_format = number_format
        self.protection = protection
        self.underline = underline
        self.border_type = border_type
        self.horizontal_alignment = horizontal_alignment
        self.vertical_alignment = vertical_alignment
        self.bg_color = get_color_from_string(bg_color, default_color=utils.colors.white)
        self.font_color = get_color_from_string(font_color, default_color=utils.colors.black)
        self.shrink_to_fit = shrink_to_fit
        self.wrap_text = wrap_text
        self.fill_pattern_type = fill_pattern_type
        self.indent = indent
        self.comment_author = comment_author
        self.comment_text = comment_text

    def __eq__(self, other):
        if not isinstance(other, self.__class__):
            return False
        return self.__dict__ == other.__dict__

    def __hash__(self):
        return hash(tuple((k, v) for k, v in self.__dict__.items()))

    def __add__(self, other):
        default = Styler().__dict__
        d = dict(self.__dict__)
        for k, v in other.__dict__.items():
            if v != default[k]:
                d[k] = v
        return Styler(**d)

    def __repr__(self):
        return pformat(self.__dict__)

    def generate_comment(self):
        if any((self.comment_author, self.comment_text)):
            return Comment(self.comment_text, self.comment_author)
        return None

    @classmethod
    def default_header_style(cls):
        return cls(bold=True)

    def to_openpyxl_style(self):
        name = str(self)
        try:
            openpyxl_style = self.cache[name]
        except KeyError:
            side = Side(border_style=self.border_type, color=utils.colors.black)
            border = Border(left=side, right=side, top=side, bottom=side)
            openpyxl_style = self.cache[name] = NamedStyle(name=name,
                                                           font=Font(name=self.font, size=self.font_size,
                                                                     color=OpenPyColor(self.font_color),
                                                                     bold=self.bold,
                                                                     underline=self.underline),
                                                           fill=PatternFill(patternType=self.fill_pattern_type,
                                                                            fgColor=self.bg_color),
                                                           alignment=Alignment(horizontal=self.horizontal_alignment,
                                                                               vertical=self.vertical_alignment,
                                                                               wrap_text=self.wrap_text,
                                                                               shrink_to_fit=self.shrink_to_fit,
                                                                               indent=self.indent),
                                                           border=border,
                                                           number_format=self.number_format,
                                                           protection=Protection(locked=self.protection))
        return openpyxl_style

    @classmethod
    def from_openpyxl_style(cls, openpyxl_style, theme_colors, openpyxl_comment=None):
        def _calc_new_hex_from_theme_hex_and_tint(theme_hex, color_tint):
            if not theme_hex.startswith('#'):
                theme_hex = '#' + theme_hex
            color_obj = Color(theme_hex)
            color_obj.luminance = _calc_lum_from_tint(color_tint, color_obj.luminance)
            return color_obj.hex_l[1:]

        def _calc_lum_from_tint(color_tint, current_lum):
            # based on http://ciintelligence.blogspot.co.il/2012/02/converting-excel-theme-color-and-tint.html
            if not color_tint:
                return current_lum
            return current_lum * (1.0 + color_tint)

        if isinstance(openpyxl_style, NamedStyle):
            openpyxl_style = openpyxl_style.name
        style_json = utils.style_str_to_dict(openpyxl_style)
        bg_color = style_json["bg_color"]

        # in case we are dealing with a "theme color"
        if not isinstance(bg_color, str):
            raise NotImplementedError("Themes not implemented yet.")
            # try:
            #     bg_color = theme_colors[openpyxl_style.fill.fgColor.theme]
            # except (AttributeError, IndexError, TypeError):
            #     bg_color = utils.colors.white[:6]
            # tint = openpyxl_style.fill.fgColor.tint
            # bg_color = _calc_new_hex_from_theme_hex_and_tint(bg_color, tint)

        bold = style_json["bold"]
        font = style_json["font"]
        font_size = style_json["font_size"]
        font_color = style_json["font_color"]

        # in case we are dealing with a "theme color"
        if not isinstance(font_color, str):
            raise NotImplementedError("Themes not implemented yet.")
            # try:
            #     font_color = theme_colors[openpyxl_style.font.color.theme]
            # except (AttributeError, IndexError, TypeError):
            #     font_color = utils.colors.black[:6]
            # tint = openpyxl_style.font.color.tint
            # font_color = _calc_new_hex_from_theme_hex_and_tint(font_color, tint)

        number_format = style_json["number_format"]
        protection = style_json["protection"]
        underline = style_json["underline"]
        border_type = style_json["border_type"]
        horizontal_alignment = style_json["horizontal_alignment"]
        vertical_alignment = style_json["vertical_alignment"]
        wrap_text = style_json["wrap_text"]
        shrink_to_fit = style_json["shrink_to_fit"]
        fill_pattern_type = style_json["fill_pattern_type"]
        indent = style_json["indent"]

        if openpyxl_comment:
            comment_author = openpyxl_comment.author
            comment_text = openpyxl_comment.text
        else:
            comment_author = style_json["comment_author"]
            comment_text = style_json["comment_text"]

        return cls(bg_color, bold, font, font_size, font_color,
                   number_format, protection, underline,
                   border_type, horizontal_alignment,
                   vertical_alignment, wrap_text, shrink_to_fit,
                   fill_pattern_type, indent, comment_author, comment_text)

    @classmethod
    def combine(cls, *styles):
        return sum(styles, cls())

    create_style = to_openpyxl_style


class ColorScaleConditionalFormatRule(object):
    """Creates a color scale conditional format rule. Wraps openpyxl's ColorScaleRule.
    Mostly should not be used directly, but through StyleFrame.add_color_scale_conditional_formatting
    """

    def __init__(self, start_type, start_value, start_color, end_type, end_value, end_color,
                 mid_type=None, mid_value=None, mid_color=None, columns_range=None):

        self.columns = columns_range

        # checking against None explicitly since mid_value may be 0
        if all(val is not None for val in (mid_type, mid_value, mid_color)):
            self.rule = ColorScaleRule(start_type=start_type, start_value=start_value,
                                       start_color=OpenPyColor(start_color),
                                       mid_type=mid_type, mid_value=mid_value,
                                       mid_color=OpenPyColor(mid_color),
                                       end_type=end_type, end_value=end_value,
                                       end_color=OpenPyColor(end_color))
        else:
            self.rule = ColorScaleRule(start_type=start_type, start_value=start_value,
                                       start_color=OpenPyColor(start_color),
                                       end_type=end_type, end_value=end_value,
                                       end_color=OpenPyColor(end_color))
