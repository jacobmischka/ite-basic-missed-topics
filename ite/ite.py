#!/usr/bin/env python3

import sys, re

from ite_section import IteSection
from ite_excel import dump_section_xlsx

def extract(inpath):
	headings = []
	ns = []
	rows = []

	with open(inpath, 'r') as infile:
		for line in infile:
			if should_skip(line):
				continue

			if is_heading_line(line):
				if not headings:
					headings = extract_headings(line)
			elif is_n_line(line):
				if not ns:
					ns = extract_ns(line)
			elif is_data_line(line):
				try:
					rows.append(extract_data_row(line))
				except AttributeError as e:
					print('Could not append row, skipping: {}'.format(e), file=sys.stderr)
			else:
				# Probably a new section
				rows.append([line.strip()])

	return headings, ns, rows

def extract_data_row(line):
	line = line.strip()
	match = re.search(r'(\(A\))|(\(B\))', line)
	keyword = line[:match.end()]
	pieces = line[match.end() + 1:]
	return [keyword, *pieces.split(' ')]

def extract_headings(line):
	return ['{}{}'.format(
		'% ' if i > 0 else '',
		heading.strip()
	) for i, heading in enumerate(line.strip().split('%'))]

def extract_ns(line):
	return [n.replace('N=', '') for n in line.strip().split(' ')]

def is_data_line(line):
	return '(A)' in line or '(B)' in line

def is_heading_line(line):
	return 'Keyword' in line

def is_n_line(line):
	return 'N=' in line

def should_skip(line):
	return (
		not line
		or len(line.strip()) == 0
		or 'Page' in line
	)

def extract_sections(rows):
	sections = []
	heading = None
	subheading = None
	items = []

	for row in rows:
		if len(row) == 1:
			if not heading:
				heading = row[0]
			elif not subheading:
				subheading = row[0]
			elif items:
				try:
					sections.append(IteSection(heading, subheading, items))
					heading = row[0]
					subheading = None
					items = []
				except Exception as e:
					print(e, file=sys.stderr)
		else:
			items.append(row)

	try:
		sections.append(IteSection(heading, subheading, items))
	except Exception as e:
		print(e, file=sys.stderr)

	return sections

def main():
	_, _, body = extract('/home/mischka/Downloads/ite and basic stuff/hm.txt')
	sections = extract_sections(body)

	# print(headings, ns, body, sep='\n\n')
	print(sections)

	dump_section_xlsx(sections, './output/2017-ite-sections.xlsx')


if __name__ == '__main__':
	main()
