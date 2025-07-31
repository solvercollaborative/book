import os
import shutil
from zipfile import ZipFile
from lxml import etree
import unittest
import tempfile

def create_simple_docx_with_bookmark(output_path):
    """Create a simple DOCX file with one point bookmark for testing."""
    NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    
    # Create minimal DOCX structure
    temp_dir = "temp_create_bookmark"
    
    try:
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)
        os.makedirs(temp_dir)
        os.makedirs(os.path.join(temp_dir, 'word'))
        os.makedirs(os.path.join(temp_dir, '_rels'))
        
        # Create [Content_Types].xml
        content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>'''
        with open(os.path.join(temp_dir, '[Content_Types].xml'), 'w', encoding='utf-8') as f:
            f.write(content_types)
        
        # Create _rels/.rels
        rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''
        with open(os.path.join(temp_dir, '_rels', '.rels'), 'w', encoding='utf-8') as f:
            f.write(rels)
        
        # Create word/document.xml with a point bookmark
        document_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{NS['w']}">
    <w:body>
        <w:p>
            <w:r>
                <w:t>Hello </w:t>
            </w:r>
            <w:bookmarkStart w:id="0" w:name="TestBookmark"/>
            <w:bookmarkEnd w:id="0"/>
            <w:r>
                <w:t>World!</w:t>
            </w:r>
        </w:p>
    </w:body>
</w:document>'''
        with open(os.path.join(temp_dir, 'word', 'document.xml'), 'w', encoding='utf-8') as f:
            f.write(document_xml)
        
        # Package into DOCX
        import zipfile
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as docx_zip:
            for foldername, subfolders, filenames in os.walk(temp_dir):
                for filename in filenames:
                    file_path = os.path.join(foldername, filename)
                    arcname = os.path.relpath(file_path, temp_dir)
                    docx_zip.write(file_path, arcname)
        
    finally:
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)

def count_point_bookmarks(docx_path):
    """Count the instances of point bookmarks in a DOCX file."""
    NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    temp_dir = "temp_count_test_bookmarks"
    
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
    test_file = "test_bookmark.docx"
    
    print("Creating test DOCX file with one point bookmark...")
    create_simple_docx_with_bookmark(test_file)
    
    print(f"Counting point bookmarks in '{test_file}'...")
    bookmark_count = count_point_bookmarks(test_file)
    print(f"Number of point bookmarks found: {bookmark_count}")
    
    # Clean up test file
    if os.path.exists(test_file):
        os.remove(test_file)
    
    return bookmark_count

class TestBookmarkCreation(unittest.TestCase):
    def test_create_and_count_bookmark(self):
        """Test that we can create a point bookmark and count it correctly."""
        test_file = "test_bookmark_unit.docx"
        
        try:
            # Create test file with one bookmark
            create_simple_docx_with_bookmark(test_file)
            
            # Count bookmarks
            actual_count = count_point_bookmarks(test_file)
            expected_count = 1
            
            self.assertEqual(actual_count, expected_count, 
                            f"Expected {expected_count} point bookmark, but found {actual_count}")
        
        finally:
            # Clean up
            if os.path.exists(test_file):
                os.remove(test_file)

if __name__ == "__main__":
    # Run the main program
    print("=== Bookmark Creation Test ===")
    result_count = main()
    
    print("\n=== Running Unit Test ===")
    # Run the unit test
    unittest.main(argv=[''], exit=False, verbosity=2) 