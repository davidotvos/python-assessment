#!/usr/bin/env python3
import json
import argparse


def main():
    # Instantiate the parser
    parser = argparse.ArgumentParser(description='Generate a pptx report from a configuration file.')
    parser.add_argument('config_file', help='Path to the JSON configuration file')

    args = parser.parse_args()
##############################################################################

if __name__ == "__main__":
    main()