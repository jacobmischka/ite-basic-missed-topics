#!/usr/bin/env python3

from argparse import ArgumentParser

def fix_higher_percentages(inpath, outpath):
	percentage_line = None

	with open(inpath, 'r') as infile, open(outpath, 'w') as outfile:
		for line in infile:
			split = line.split(' ')
			if len(split) == 2 and '%' in split[0] and '%' in split[1]:
				percentage_line = line
			else:
				if percentage_line:
					outfile.write(' '.join([line.strip(), percentage_line]))
				else:
					outfile.write(line)

def main():
	parser = ArgumentParser()
	parser.add_argument('inpath')
	parser.add_argument('outpath')

	args = parser.parse_args()

	fix_higher_percentages(args.inpath, args.outpath)

if __name__ == '__main__':
	main()
