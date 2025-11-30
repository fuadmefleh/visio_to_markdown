#!/usr/bin/env python3
"""
Test suite for the Visio to Markdown converter.

Run with:
    python test_visio_to_markdown.py
    python -m pytest test_visio_to_markdown.py -v
"""

import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, MagicMock, patch

from visio_to_markdown_standalone import VisioToMarkdownConverter


class TestVisioToMarkdownConverter(unittest.TestCase):
    """Test cases for VisioToMarkdownConverter class."""
    
    def setUp(self):
        """Set up test fixtures."""
        self.converter = VisioToMarkdownConverter(verbose=False)
    
    def test_initialization(self):
        """Test converter initialization."""
        converter = VisioToMarkdownConverter(verbose=True)
        self.assertIsNotNone(converter)
        self.assertIsNotNone(converter.logger)
    
    def test_sanitize_mermaid_id(self):
        """Test Mermaid ID sanitization."""
        # Test with normal text
        result = self.converter._sanitize_mermaid_id("Hello World")
        self.assertEqual(result, "Hello_World")
        
        # Test with special characters
        result = self.converter._sanitize_mermaid_id("Test-123!@#$%")
        self.assertEqual(result, "Test_123_____")
        
        # Test with long text (should be truncated to 50 chars)
        long_text = "A" * 100
        result = self.converter._sanitize_mermaid_id(long_text)
        self.assertEqual(len(result), 50)
        
        # Test with shape_id fallback
        result = self.converter._sanitize_mermaid_id("", shape_id=123)
        self.assertEqual(result, "shape_123")
        
        # Test with no text or ID
        result = self.converter._sanitize_mermaid_id("")
        self.assertEqual(result, "unknown")
    
    def test_get_attribute_value(self):
        """Test safe attribute value extraction."""
        # Test with property
        obj = Mock()
        obj.test_attr = "test_value"
        result = self.converter._get_attribute_value(obj, 'test_attr')
        self.assertEqual(result, "test_value")
        
        # Test with method
        obj = Mock()
        obj.test_method = Mock(return_value="method_value")
        result = self.converter._get_attribute_value(obj, 'test_method')
        self.assertEqual(result, "method_value")
        
        # Test with non-existent attribute
        result = self.converter._get_attribute_value(obj, 'nonexistent', default='default')
        self.assertEqual(result, 'default')
        
        # Test with exception during access
        obj = Mock()
        obj.bad_attr = Mock(side_effect=Exception("Error"))
        result = self.converter._get_attribute_value(obj, 'bad_attr', default='fallback')
        self.assertEqual(result, 'fallback')
    
    def test_get_shapes_from_object(self):
        """Test shape extraction from objects."""
        # Test with shapes as list property
        obj = Mock()
        shape1 = Mock()
        shape2 = Mock()
        obj.shapes = [shape1, shape2]
        
        result = self.converter._get_shapes_from_object(obj)
        self.assertEqual(len(result), 2)
        self.assertIn(shape1, result)
        self.assertIn(shape2, result)
        
        # Test with shapes as method
        obj = Mock()
        obj.shapes = Mock(return_value=[shape1])
        result = self.converter._get_shapes_from_object(obj)
        self.assertEqual(len(result), 1)
        
        # Test with no shapes
        obj = Mock(spec=[])
        result = self.converter._get_shapes_from_object(obj)
        self.assertEqual(len(result), 0)
    
    def test_extract_connections(self):
        """Test connection extraction from shapes."""
        # Create mock shape with connections
        shape = Mock()
        connect = Mock()
        from_shape = Mock()
        from_shape.ID = "1"
        to_shape = Mock()
        to_shape.ID = "2"
        
        connect.from_shape = from_shape
        connect.to_shape = to_shape
        shape.connects = [connect]
        
        result = self.converter._extract_connections(shape)
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0], ("1", "2"))
        
        # Test with no connections
        shape = Mock(spec=['text'])
        result = self.converter._extract_connections(shape)
        self.assertEqual(len(result), 0)
    
    def test_extract_shape_info(self):
        """Test shape information extraction."""
        # Create a mock shape
        shape = Mock()
        shape.text = "Test Shape"
        shape.name = "Shape1"
        shape.ID = 123
        shape.id = 123
        
        # Mock master shape
        master = Mock()
        master.name = "Rectangle"
        shape.master_shape = master
        
        # Mock no sub-shapes
        shape.shapes = []
        shape.child_shapes = []
        shape.sub_shapes = []
        
        result = self.converter._extract_shape_info(shape, depth=0)
        
        self.assertEqual(result['text'], "Test Shape")
        self.assertEqual(result['name'], "Shape1")
        self.assertEqual(result['id'], 123)
        self.assertEqual(result['type'], "Rectangle")
        self.assertFalse(result['has_image'])
        self.assertEqual(len(result['sub_shapes']), 0)
    
    def test_extract_shape_info_with_depth_limit(self):
        """Test that shape extraction respects depth limit."""
        # Create nested shapes
        shape = Mock()
        shape.text = "Parent"
        shape.ID = 1
        
        sub_shape = Mock()
        sub_shape.text = "Child"
        sub_shape.ID = 2
        sub_shape.shapes = []
        
        shape.shapes = [sub_shape]
        
        # Extract at max depth - should not process sub-shapes
        result = self.converter._extract_shape_info(shape, depth=5)
        self.assertEqual(len(result['sub_shapes']), 0)
        
        # Extract at normal depth - should process sub-shapes
        result = self.converter._extract_shape_info(shape, depth=0)
        self.assertEqual(len(result['sub_shapes']), 1)
    
    def test_generate_mermaid_diagram_basic(self):
        """Test Mermaid diagram generation."""
        page_data = {
            "shapes": [
                {"id": "1", "text": "Start", "has_image": False},
                {"id": "2", "text": "Process", "has_image": False},
                {"id": "3", "text": "End", "has_image": False}
            ],
            "connections": [("1", "2"), ("2", "3")]
        }
        
        result = self.converter._generate_mermaid_diagram(page_data)
        
        self.assertIn("```mermaid", result)
        self.assertIn("graph TD", result)
        self.assertIn("Start", result)
        self.assertIn("Process", result)
        self.assertIn("End", result)
        self.assertIn("-->", result)
        self.assertIn("```", result)
    
    def test_generate_mermaid_diagram_with_images(self):
        """Test Mermaid diagram with image markers."""
        page_data = {
            "shapes": [
                {"id": "1", "text": "Image Shape", "has_image": True}
            ],
            "connections": []
        }
        
        result = self.converter._generate_mermaid_diagram(page_data)
        
        self.assertIn("📷", result)
    
    def test_generate_mermaid_diagram_with_page_media(self):
        """Test Mermaid diagram with page-level media."""
        page_data = {
            "shapes": [],
            "connections": [],
            "page_media": [
                {"index": 0, "data": b"fake_image_data"}
            ]
        }
        
        result = self.converter._generate_mermaid_diagram(page_data)
        
        self.assertIn("page_image_0", result)
        self.assertIn("📷 Page Image 1", result)
    
    def test_to_markdown_basic(self):
        """Test basic markdown conversion."""
        data = {
            "file_name": "test.vsdx",
            "pages": [
                {
                    "name": "Page 1",
                    "shapes": [
                        {"id": "1", "text": "Test Shape", "name": "Shape1", 
                         "type": "Rectangle", "has_image": False, "sub_shapes": []}
                    ],
                    "connections": [],
                    "images_count": 0,
                    "page_media": []
                }
            ],
            "metadata": {
                "title": "Test Diagram",
                "creator": "Test User",
                "company": "Test Company"
            },
            "total_images": 0
        }
        
        result = self.converter._to_markdown(data)
        
        self.assertIn("# test.vsdx", result)
        self.assertIn("## Metadata", result)
        self.assertIn("Test Diagram", result)
        self.assertIn("Test User", result)
        self.assertIn("Test Company", result)
        self.assertIn("Page 1", result)
        self.assertIn("Test Shape", result)
    
    def test_to_markdown_with_connections(self):
        """Test markdown with connection information."""
        data = {
            "file_name": "test.vsdx",
            "pages": [
                {
                    "name": "Page 1",
                    "shapes": [
                        {"id": "1", "text": "Shape1", "sub_shapes": [], "has_image": False},
                        {"id": "2", "text": "Shape2", "sub_shapes": [], "has_image": False}
                    ],
                    "connections": [("1", "2")],
                    "images_count": 0,
                    "page_media": []
                }
            ],
            "metadata": {},
            "total_images": 0
        }
        
        result = self.converter._to_markdown(data)
        
        self.assertIn("Connections", result)
        self.assertIn("Shape 1 → Shape 2", result)
    
    def test_convert_invalid_file(self):
        """Test conversion with invalid file path."""
        with self.assertRaises(FileNotFoundError):
            self.converter.convert("/nonexistent/file.vsdx")
    
    def test_convert_invalid_format(self):
        """Test conversion with invalid file format."""
        with tempfile.NamedTemporaryFile(suffix='.txt', delete=False) as tmp:
            tmp_path = tmp.name
        
        try:
            with self.assertRaises(ValueError) as context:
                self.converter.convert(tmp_path)
            self.assertIn("Unsupported file format", str(context.exception))
        finally:
            Path(tmp_path).unlink()
    
    @patch('visio_to_markdown_standalone.VisioFile')
    def test_convert_markdown_format(self, mock_visio_file):
        """Test conversion to markdown format."""
        # Create a temporary vsdx file
        with tempfile.NamedTemporaryFile(suffix='.vsdx', delete=False) as tmp:
            tmp_path = tmp.name
        
        try:
            # Mock the VisioFile context manager
            mock_vis = MagicMock()
            mock_page = Mock()
            mock_page.name = "Test Page"
            mock_page.shapes = []
            
            mock_vis.pages = [mock_page]
            mock_vis.__enter__ = Mock(return_value=mock_vis)
            mock_vis.__exit__ = Mock(return_value=False)
            
            mock_visio_file.return_value = mock_vis
            
            result = self.converter.convert(tmp_path, output_format='markdown')
            
            self.assertIsInstance(result, str)
            self.assertIn("# ", result)
        finally:
            Path(tmp_path).unlink()
    
    @patch('visio_to_markdown_standalone.VisioFile')
    def test_convert_json_format(self, mock_visio_file):
        """Test conversion to JSON format."""
        with tempfile.NamedTemporaryFile(suffix='.vsdx', delete=False) as tmp:
            tmp_path = tmp.name
        
        try:
            # Mock the VisioFile
            mock_vis = MagicMock()
            mock_page = Mock()
            mock_page.name = "Test Page"
            mock_page.shapes = []
            
            mock_vis.pages = [mock_page]
            mock_vis.__enter__ = Mock(return_value=mock_vis)
            mock_vis.__exit__ = Mock(return_value=False)
            
            mock_visio_file.return_value = mock_vis
            
            result = self.converter.convert(tmp_path, output_format='json')
            
            self.assertIsInstance(result, dict)
            self.assertIn('file_name', result)
            self.assertIn('pages', result)
        finally:
            Path(tmp_path).unlink()
    
    @patch('visio_to_markdown_standalone.VisioFile')
    def test_convert_both_formats(self, mock_visio_file):
        """Test conversion to both formats."""
        with tempfile.NamedTemporaryFile(suffix='.vsdx', delete=False) as tmp:
            tmp_path = tmp.name
        
        try:
            # Mock the VisioFile
            mock_vis = MagicMock()
            mock_page = Mock()
            mock_page.name = "Test Page"
            mock_page.shapes = []
            
            mock_vis.pages = [mock_page]
            mock_vis.__enter__ = Mock(return_value=mock_vis)
            mock_vis.__exit__ = Mock(return_value=False)
            
            mock_visio_file.return_value = mock_vis
            
            result = self.converter.convert(tmp_path, output_format='both')
            
            self.assertIsInstance(result, tuple)
            self.assertEqual(len(result), 2)
            
            markdown, json_data = result
            self.assertIsInstance(markdown, str)
            self.assertIsInstance(json_data, dict)
        finally:
            Path(tmp_path).unlink()
    
    def test_extract_page_data(self):
        """Test page data extraction."""
        # Create a mock page
        mock_page = Mock()
        mock_page.name = "Test Page"
        
        # Mock shapes
        mock_shape = Mock()
        mock_shape.text = "Test Shape"
        mock_shape.ID = 1
        mock_shape.name = "Shape1"
        mock_shape.shapes = []
        
        mock_page.shapes = [mock_shape]
        
        result = self.converter._extract_page_data(mock_page)
        
        self.assertEqual(result['name'], "Test Page")
        self.assertIsInstance(result['shapes'], list)
        self.assertEqual(len(result['shapes']), 1)
        self.assertEqual(result['images_count'], 0)


class TestCLIIntegration(unittest.TestCase):
    """Test cases for command-line interface."""
    
    @patch('visio_to_markdown_standalone.VisioToMarkdownConverter')
    @patch('sys.argv', ['visio_to_markdown_standalone.py', 'test.vsdx'])
    def test_main_basic(self, mock_converter_class):
        """Test basic CLI execution."""
        from visio_to_markdown_standalone import main
        
        # Mock converter
        mock_converter = Mock()
        mock_converter.convert = Mock(return_value="# Test Output")
        mock_converter_class.return_value = mock_converter
        
        # Create temp file
        with tempfile.NamedTemporaryFile(suffix='.vsdx', delete=False) as tmp:
            tmp_path = tmp.name
        
        try:
            with patch('sys.argv', ['script.py', tmp_path]):
                with patch('builtins.print') as mock_print:
                    result = main()
                    self.assertEqual(result, 0)
                    mock_converter.convert.assert_called_once()
        finally:
            Path(tmp_path).unlink()
    
    def test_main_with_output_file(self):
        """Test CLI with output file."""
        from visio_to_markdown_standalone import main
        
        with tempfile.NamedTemporaryFile(suffix='.vsdx', delete=False) as input_tmp:
            input_path = input_tmp.name
        
        with tempfile.NamedTemporaryFile(suffix='.md', delete=False) as output_tmp:
            output_path = output_tmp.name
        
        try:
            with patch('visio_to_markdown_standalone.VisioToMarkdownConverter') as mock_class:
                mock_converter = Mock()
                mock_converter.convert = Mock(return_value="# Test")
                mock_class.return_value = mock_converter
                
                with patch('sys.argv', ['script.py', input_path, '-o', output_path]):
                    with patch('builtins.print'):
                        result = main()
                        self.assertEqual(result, 0)
        finally:
            Path(input_path).unlink()
            if Path(output_path).exists():
                Path(output_path).unlink()


class TestEdgeCases(unittest.TestCase):
    """Test edge cases and error handling."""
    
    def setUp(self):
        self.converter = VisioToMarkdownConverter(verbose=False)
    
    def test_empty_page_data(self):
        """Test handling of empty page data."""
        data = {
            "file_name": "empty.vsdx",
            "pages": [],
            "metadata": {},
            "total_images": 0
        }
        
        result = self.converter._to_markdown(data)
        self.assertIn("# empty.vsdx", result)
        self.assertIn("0 total", result)
    
    def test_shape_with_no_text(self):
        """Test handling of shapes without text."""
        shape = Mock()
        shape.text = ""
        shape.name = ""
        shape.ID = None
        shape.shapes = []
        
        result = self.converter._extract_shape_info(shape)
        
        self.assertEqual(result['text'], "")
        self.assertEqual(result['name'], "")
    
    def test_mermaid_special_characters(self):
        """Test Mermaid ID with special characters."""
        page_data = {
            "shapes": [
                {"id": "1", "text": 'Shape with "quotes"', "has_image": False}
            ],
            "connections": []
        }
        
        result = self.converter._generate_mermaid_diagram(page_data)
        
        # Quotes should be replaced with single quotes
        self.assertIn("'", result)
        self.assertNotIn('Shape with "quotes"', result)
    
    def test_metadata_with_empty_values(self):
        """Test metadata handling with empty values."""
        data = {
            "file_name": "test.vsdx",
            "pages": [],
            "metadata": {
                "title": "",
                "creator": None,
                "company": ""
            },
            "total_images": 0
        }
        
        result = self.converter._to_markdown(data)
        
        # Metadata section should not appear if all values are empty
        self.assertNotIn("## Metadata", result)


def run_tests():
    """Run all tests."""
    # Create test suite
    loader = unittest.TestLoader()
    suite = unittest.TestSuite()
    
    # Add all test cases
    suite.addTests(loader.loadTestsFromTestCase(TestVisioToMarkdownConverter))
    suite.addTests(loader.loadTestsFromTestCase(TestCLIIntegration))
    suite.addTests(loader.loadTestsFromTestCase(TestEdgeCases))
    
    # Run tests
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    
    # Return exit code
    return 0 if result.wasSuccessful() else 1


if __name__ == '__main__':
    exit(run_tests())
