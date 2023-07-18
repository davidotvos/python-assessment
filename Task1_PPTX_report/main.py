#!/usr/bin/env python3
import json
import argparse
import os

# 
def validate_json_file(file_path):
    if not file_path.lower().endswith('.json') or not os.path.isfile(file_path):
        raise argparse.ArgumentTypeError('Invalid JSON file: {}'.format(file_path))
    return file_path

def main():
    # Instantiate the parser
    parser = argparse.ArgumentParser(description='Generate a pptx report from a configuration file.')
    parser.add_argument('config_file',type=validate_json_file, help='Path to the JSON configuration file')

    args = parser.parse_args()
    config_file = args.config_file

    print(config_file)
##############################################################################

if __name__ == "__main__":
    main()