#!/usr/bin/env python3

import sys
import zipfile
from lxml import etree
import os
import shutil
import unittest

# Use command line argument if provided, otherwise use default
input_file = sys.argv[1] if len(sys.argv) > 1 else "../mybooks/superArchItelligence Vol1 8x10.docx"

def count_xe_tags(docx_path):
    """Count the instances of XE tags in a DOCX file."""
    NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    temp_dir = "temp_count_xe"
    
    try:
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)
        
        with zipfile.ZipFile(docx_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
        
        xml_path = os.path.join(temp_dir, 'word', 'document.xml')
        parser = etree.XMLParser(ns_clean=True)
        
        with open(xml_path, 'rb') as f:
            tree = etree.parse(f, parser)
        
        root = tree.getroot()
        
        # Find all instrText nodes containing XE tags
        instr_texts = root.xpath('//w:instrText[contains(text(), "XE")]', namespaces=NS)
        
        xe_count = 0
        for instr in instr_texts:
            field_text = instr.text or ""
            # Case-sensitive search for 'XE "'
            xe_count += field_text.count('XE "')
        
        return xe_count
    
    finally:
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)

def main():
    if not os.path.exists(input_file):
        print(f"Error: Input file not found: {input_file}")
        return
    
    xe_count = count_xe_tags(input_file)
    print(f"Number of XE tags found in '{input_file}': {xe_count}")
    
    return xe_count

class TestXETagCount(unittest.TestCase):
    def test_xe_tag_count(self):
        """Test that the XE tag count is exactly 402."""
        actual_count = count_xe_tags(input_file)
        expected_count = 402
        self.assertEqual(actual_count, expected_count, 
                        f"Expected {expected_count} XE tags, but found {actual_count}")

if __name__ == "__main__":
    # Run the main program
    print("=== XE Tag Counter ===")
    result_count = main()
    
    print("\n=== Running Unit Test ===")
    # Run the unit test
    unittest.main(argv=[''], exit=False, verbosity=2) 