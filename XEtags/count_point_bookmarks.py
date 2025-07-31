#!/usr/bin/env python3

import sys
import os
import shutil
from zipfile import ZipFile
from lxml import etree
import unittest

# Use command line argument if provided, otherwise use default
INPUT_FILE = sys.argv[1] if len(sys.argv) > 1 else "../mybooks/superArchItelligence Vol1 8x10.docm"

def count_point_bookmarks(docx_path):
    """Count the instances of point bookmarks in a DOCX file."""
    NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    temp_dir = "temp_count_bookmarks"
    
    try:
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)
        
        with ZipFile(docx_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
        
        xml_path = os.path.join(temp_dir, 'word', 'document.xml')
        parser = etree.XMLParser(ns_clean=True)
        
        with open(xml_path, 'rb') as f:
            tree = etree.parse(f, parser)
        
        root = tree.getroot()
        
        # Find all bookmarkStart elements
        bookmark_starts = root.xpath('//w:bookmarkStart', namespaces=NS)
        
        point_bookmark_count = 0
        for bookmark_start in bookmark_starts:
            bookmark_id = bookmark_start.get(f'{{{NS["w"]}}}id')
            if bookmark_id:
                # Find the corresponding bookmarkEnd with the same id
                bookmark_end = root.xpath(f'//w:bookmarkEnd[@w:id="{bookmark_id}"]', namespaces=NS)
                if bookmark_end:
                    # Check if this is a point bookmark (start and end are adjacent)
                    parent = bookmark_start.getparent()
                    if parent is not None:
                        children = list(parent)
                        start_index = children.index(bookmark_start)
                        # Check if the next element is the corresponding bookmarkEnd
                        if (start_index + 1 < len(children) and 
                            children[start_index + 1].tag == f'{{{NS["w"]}}}bookmarkEnd' and
                            children[start_index + 1].get(f'{{{NS["w"]}}}id') == bookmark_id):
                            point_bookmark_count += 1
        
        return point_bookmark_count
    
    finally:
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)

def main():
    if not os.path.exists(INPUT_FILE):
        print(f"Error: Input file not found: {INPUT_FILE}")
        return
    
    bookmark_count = count_point_bookmarks(INPUT_FILE)
    print(f"Number of point bookmarks found in '{INPUT_FILE}': {bookmark_count}")
    
    return bookmark_count

class TestPointBookmarkCount(unittest.TestCase):
    def test_point_bookmark_count(self):
        """Test that the point bookmark count is as expected."""
        actual_count = count_point_bookmarks(INPUT_FILE)
        print(f"Found {actual_count} point bookmarks in the document")
        # The .docm file contains exactly 1 point bookmark
        self.assertEqual(actual_count, 1, "Point bookmark count should be one")

if __name__ == "__main__":
    # Run the main program
    print("=== Point Bookmark Counter ===")
    result_count = main()
    
    print("\n=== Running Unit Test ===")
    # Run the unit test
    unittest.main(argv=[''], exit=False, verbosity=2)