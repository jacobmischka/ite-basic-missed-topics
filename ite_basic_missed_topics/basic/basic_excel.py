from openpyxl import Workbook
from openpyxl.styles import NamedStyle, Font, PatternFill, Alignment
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils.cell import get_column_letter

from .basic_section import BasicSection, BasicItem
from ..utils import (
	get_data_ranges
)
from ..constants import (
	DATA_START_ROW
)

def dump_section_xlsx(sections, outpath):
	wb = Workbook()
	ws = wb.active

	row = DATA_START_ROW
	row_ranges = []

	for section in sections:
		next_row = section.write_xlsx_rows(ws, row)
		row_ranges.append((
			row,
			next_row
		))
		row = next_row

	data_ranges = get_data_ranges(row_ranges)
	add_styles(ws, data_ranges)

	wb.save(outpath)


def add_styles(worksheet, data_ranges):
	subheading = NamedStyle(name='subheading')
	subheading.font = Font(name='Calibri', size=11, color='ffffff')
	subheading.fill = PatternFill(start_color='444444', end_color='444444',
		fill_type='solid')
	subheading.number_format = '0%'

	for range_start, _ in data_ranges:
		for row in worksheet['{}{}:{}{}'.format(
			get_column_letter(BasicItem.KEYWORD_COL),
			range_start - 2,
			get_column_letter(BasicItem.MISSED_COL),
			range_start - 1
		)]:
			for cell in row:
				cell.style = subheading


	for col in BasicItem.SEPARATOR_COLS:
		for row in worksheet.iter_rows(min_row=DATA_START_ROW):
			row[col - 1].style = subheading

# def add_conditional_formatting(worksheet, data_ranges):
# 	diff_range =
