import os
import shutil
from zipfile import ZipFile
from lxml import etree
import unittest
import re

# Define input and output file paths
INPUT_FILE = "../mybooks/superArchItelligence Vol1 8x10.docm"
OUTPUT_FILE = "../mybooks/superArchItelligence Vol1 e-book.docm"

def sanitize_bookmark_name(name):
    """Bookmark names must start with a letter and contain only letters, numbers, and underscores"""
    name = re.sub(r'[^\w]', '_', name)
    if not name[0].isalpha():
        name = 'B_' + name
    return name[:40]  # Bookmark names must be ≤ 40 characters

def repackage_docx_from_dir(temp_dir, output_path):
    """Properly repackage a DOCX directory structure into a zip file."""
    import zipfile
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as docx_zip:
        for foldername, subfolders, filenames in os.walk(temp_dir):
            for filename in filenames:
                file_path = os.path.join(foldername, filename)
                arcname = os.path.relpath(file_path, temp_dir)
                docx_zip.write(file_path, arcname)

def convert_xe_tags_to_bookmarks(input_path, output_path):
    """Convert XE tags to point bookmarks in a DOCX/DOCM file."""
    NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    temp_dir = "temp_convert_xe"
    
    try:
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)
        
        with ZipFile(input_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
        
        xml_path = os.path.join(temp_dir, 'word', 'document.xml')
        parser = etree.XMLParser(ns_clean=True)
        
        with open(xml_path, 'rb') as f:
            tree = etree.parse(f, parser)
        
        root = tree.getroot()
        bookmark_id = 0
        
        # Find all XE-related field parts and collect them for deletion
        xe_runs_to_delete = set()
        
        # Find all instrText elements containing "XE" (complete or partial)
        instr_texts = root.xpath('//w:instrText[contains(text(), "XE")]', namespaces=NS)
        
        for instr in instr_texts:
            field_text = instr.text or ""
            
            # Create a bookmark for this XE field (even if fragmented)
            run = instr.getparent()
            paragraph = run.getparent()
            
            # Try to find a nearby run with display text
            text_run = None
            run_index = paragraph.index(run)
            for i in reversed(range(run_index)):
                prior = paragraph[i]
                if prior.tag.endswith('r') and prior.xpath('.//w:t', namespaces=NS):
                    text_run = prior
                    break
            
            if text_run is not None:
                # Create bookmark with sequential name
                bookmark_name = f"xe_bookmark_{bookmark_id}"
                bookmark_start = etree.Element('{%s}bookmarkStart' % NS['w'])
                bookmark_start.set('{%s}id' % NS['w'], str(bookmark_id))
                bookmark_start.set('{%s}name' % NS['w'], bookmark_name)
                
                bookmark_end = etree.Element('{%s}bookmarkEnd' % NS['w'])
                bookmark_end.set('{%s}id' % NS['w'], str(bookmark_id))
                
                bookmark_id += 1
                
                # Create point bookmark by placing start and end adjacent to each other
                text_run.addprevious(bookmark_start)
                text_run.addprevious(bookmark_end)  # Both before the text run, making them adjacent
            
            # Mark this run for deletion
            xe_runs_to_delete.add(run)
        
        # Also find any runs that contain XE field characters or related content
        # Look for fldChar elements that might be part of XE fields
        for fld_char in root.xpath('//w:fldChar', namespaces=NS):
            run = fld_char.getparent()
            # Check if this run or nearby runs contain XE-related content
            paragraph = run.getparent()
            run_index = paragraph.index(run)
            
            # Check runs around this fldChar for XE content
            for i in range(max(0, run_index-2), min(len(paragraph), run_index+3)):
                check_run = paragraph[i]
                if check_run.tag.endswith('r'):
                    # Check for XE in instrText or regular text
                    for text_elem in check_run.xpath('.//w:instrText | .//w:t', namespaces=NS):
                        if text_elem.text and 'XE' in text_elem.text:
                            xe_runs_to_delete.add(check_run)
                            break
        
        # Delete all identified XE-related runs
        for run in xe_runs_to_delete:
            parent = run.getparent()
            if parent is not None:
                parent.remove(run)
        
        tree.write(xml_path, encoding='utf-8', xml_declaration=True)
        repackage_docx_from_dir(temp_dir, output_path)
        
    finally:
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)

def count_point_bookmarks(docx_path):
    """Count the instances of point bookmarks in a DOCX/DOCM file."""
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
    
    print(f"Converting XE tags to bookmarks...")
    print(f"Input file: {INPUT_FILE}")
    print(f"Output file: {OUTPUT_FILE}")
    
    # Convert XE tags to bookmarks
    convert_xe_tags_to_bookmarks(INPUT_FILE, OUTPUT_FILE)
    
    # Count point bookmarks in the output file
    bookmark_count = count_point_bookmarks(OUTPUT_FILE)
    print(f"Number of point bookmarks in output file: {bookmark_count}")
    
    return bookmark_count

class TestXETagConversion(unittest.TestCase):
    def test_xe_tag_conversion_count(self):
        """Test that after conversion, we have exactly 403 point bookmarks (1 original + 402 converted)."""
        expected_count = 403  # 1 valid bookmark in original + 402 XE tags converted
        actual_count = count_point_bookmarks(OUTPUT_FILE)
        self.assertEqual(actual_count, expected_count,
                        f"Expected {expected_count} point bookmarks after conversion, but found {actual_count}")

if __name__ == "__main__":
    # Run the main program
    print("=== XE Tag to Bookmark Converter ===")
    result_count = main()
    
    print("\n=== Running Unit Test ===")
    # Run the unit test
    unittest.main(argv=[''], exit=False, verbosity=2) 