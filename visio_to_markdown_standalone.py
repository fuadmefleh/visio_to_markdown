#!/usr/bin/env python3
"""
Visio to Markdown Converter

Converts Visio (.vsdx) files to Markdown format with Mermaid diagrams.
Extracts text, shapes, pages, connections, and structural information.

Usage:
    python visio_to_markdown_standalone.py input.vsdx [--output output.md] [--format markdown|json|both]

Requirements:
    pip install vsdx
"""

import argparse
import json
import logging
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union


class VisioToMarkdownConverter:
    """Converts Visio files to Markdown with Mermaid diagrams."""
    
    def __init__(self, verbose: bool = False):
        """
        Initialize the converter.
        
        Args:
            verbose: Enable verbose logging
        """
        self.logger = logging.getLogger(__name__)
        level = logging.DEBUG if verbose else logging.INFO
        logging.basicConfig(
            level=level,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )
    
    def _get_shapes_from_object(self, obj) -> List:
        """
        Safely gets shapes from an object, handling both properties and methods.
        
        Args:
            obj: Object that might have shapes
            
        Returns:
            List of shapes or empty list
        """
        shapes = []
        
        # Try different ways to get shapes
        for attr_name in ['shapes', 'child_shapes', 'sub_shapes']:
            if hasattr(obj, attr_name):
                attr = getattr(obj, attr_name)
                # Check if it's a method
                if callable(attr):
                    try:
                        result = attr()
                        if result:
                            shapes.extend(result if isinstance(result, list) else [result])
                    except:
                        pass
                # Check if it's iterable (list, tuple, etc.)
                elif attr is not None:
                    try:
                        shapes.extend(attr if isinstance(attr, list) else list(attr))
                    except:
                        pass
        
        return shapes

    def _get_attribute_value(self, obj, attr_name: str, default=None):
        """
        Safely gets an attribute value, handling both properties and methods.
        
        Args:
            obj: Object to get attribute from
            attr_name: Name of the attribute
            default: Default value if attribute doesn't exist
            
        Returns:
            Attribute value or default
        """
        try:
            if hasattr(obj, attr_name):
                attr = getattr(obj, attr_name)
                if callable(attr):
                    return attr()
                return attr
        except:
            pass
        return default

    def _extract_connections(self, shape) -> List[Tuple[str, str]]:
        """
        Extracts connection information from a shape (connectors).
        
        Args:
            shape: A vsdx shape object
            
        Returns:
            List of tuples (from_shape_id, to_shape_id)
        """
        connections = []
        
        try:
            # Check if this is a connector shape
            if hasattr(shape, 'connects'):
                connects = self._get_attribute_value(shape, 'connects', [])
                if connects:
                    from_id = None
                    to_id = None
                    
                    for connect in connects:
                        from_shape = self._get_attribute_value(connect, 'from_shape')
                        to_shape = self._get_attribute_value(connect, 'to_shape')
                        
                        if from_shape:
                            from_id = self._get_attribute_value(from_shape, 'ID')
                        if to_shape:
                            to_id = self._get_attribute_value(to_shape, 'ID')
                    
                    if from_id and to_id:
                        connections.append((str(from_id), str(to_id)))
            
            # Alternative: check for begin_x/end_x which might indicate connectors
            if hasattr(shape, 'one_d') or hasattr(shape, 'OneD'):
                # This might be a 1D shape (connector)
                begin_shape = self._get_attribute_value(shape, 'begin_shape')
                end_shape = self._get_attribute_value(shape, 'end_shape')
                
                if begin_shape and end_shape:
                    from_id = self._get_attribute_value(begin_shape, 'ID')
                    to_id = self._get_attribute_value(end_shape, 'ID')
                    
                    if from_id and to_id:
                        connections.append((str(from_id), str(to_id)))
        except:
            pass
        
        return connections

    def _extract_media_from_page(self, page) -> List[Dict[str, Any]]:
        """
        Extracts media/images directly from the page (not in shapes).
        These are typically dragged-and-dropped images.
        
        Args:
            page: A vsdx.Page object
            
        Returns:
            List of media items with their data
        """
        media_items = []
        
        try:
            # Method 1: Check for media attribute
            if hasattr(page, 'media'):
                media = self._get_attribute_value(page, 'media')
                if media:
                    self.logger.info(f"Found media attribute on page")
                    # Media might be a list or dict
                    if isinstance(media, list):
                        for idx, item in enumerate(media):
                            media_items.append({
                                'index': idx,
                                'data': item,
                                'source': 'page.media'
                            })
                    elif isinstance(media, dict):
                        for key, item in media.items():
                            media_items.append({
                                'index': key,
                                'data': item,
                                'source': 'page.media'
                            })
            
            # Method 2: Check for images attribute
            if hasattr(page, 'images'):
                images = self._get_attribute_value(page, 'images')
                if images:
                    self.logger.info(f"Found images attribute on page")
                    if isinstance(images, list):
                        for idx, item in enumerate(images):
                            media_items.append({
                                'index': idx,
                                'data': item,
                                'source': 'page.images'
                            })
            
            # Method 3: Check VisioFile directly for media files
            if hasattr(page, 'filename'):
                filename = self._get_attribute_value(page, 'filename')
                self.logger.debug(f"Page filename: {filename}")
                
        except Exception as e:
            self.logger.warning(f"Error extracting page media: {e}")
        
        return media_items

    def _extract_image_from_shape(self, shape) -> Optional[bytes]:
        """
        Extracts image data from a shape if it contains an image.
        
        Args:
            shape: A vsdx shape object
            
        Returns:
            Image bytes or None
        """
        try:
            # Log all attributes to help debug
            shape_id = self._get_attribute_value(shape, 'ID')
            if shape_id:
                all_attrs = [attr for attr in dir(shape) if not attr.startswith('_')]
                self.logger.debug(f"Shape {shape_id} attributes: {all_attrs}")
            
            # Method 1: Direct image_data attribute
            if hasattr(shape, 'image_data'):
                image_data = self._get_attribute_value(shape, 'image_data')
                if image_data:
                    self.logger.info(f"Found image_data in shape {shape_id}")
                    return image_data
            
            # Method 2: image attribute
            if hasattr(shape, 'image'):
                image = self._get_attribute_value(shape, 'image')
                if image:
                    self.logger.info(f"Found image in shape {shape_id}")
                    return image
            
            # Method 3: Check for media/file references
            if hasattr(shape, 'file'):
                file_data = self._get_attribute_value(shape, 'file')
                if file_data:
                    self.logger.info(f"Found file in shape {shape_id}")
                    return file_data
            
            # Method 4: Check master_shape for image
            if hasattr(shape, 'master_shape'):
                master = self._get_attribute_value(shape, 'master_shape')
                if master:
                    if hasattr(master, 'image_data'):
                        image_data = self._get_attribute_value(master, 'image_data')
                        if image_data:
                            self.logger.info(f"Found image_data in master_shape")
                            return image_data
            
            # Method 5: Check if shape has a fill with image
            if hasattr(shape, 'fill'):
                fill = self._get_attribute_value(shape, 'fill')
                if fill and hasattr(fill, 'image'):
                    image = self._get_attribute_value(fill, 'image')
                    if image:
                        self.logger.info(f"Found image in fill")
                        return image
            
            # Method 6: Check for cells that might contain image references
            if hasattr(shape, 'cells'):
                cells = self._get_attribute_value(shape, 'cells')
                if cells:
                    # Look for FillForegnd which might reference an image
                    for cell_name in ['FillForegnd', 'FillBkgnd', 'FillPattern']:
                        if cell_name in cells:
                            cell_value = cells[cell_name]
                            self.logger.debug(f"Found cell {cell_name}: {cell_value}")
                        
        except Exception as e:
            self.logger.warning(f"Error extracting image from shape: {e}")
        
        return None

    def _extract_shape_info(self, shape, depth: int = 0) -> Dict[str, Any]:
        """
        Recursively extracts information from a shape and its sub-shapes.
        
        Args:
            shape: A vsdx shape object
            depth: Current recursion depth
            
        Returns:
            Dict containing shape information
        """
        shape_info = {
            "text": "",
            "name": "",
            "id": None,
            "type": "",
            "sub_shapes": [],
            "connections": [],
            "has_image": False
        }
        
        # Try multiple ways to get text
        shape_info["text"] = str(self._get_attribute_value(shape, 'text', '')).strip()
        
        # Try to get shape name
        shape_info["name"] = str(self._get_attribute_value(shape, 'name', ''))
        
        # Try to get ID
        shape_info["id"] = self._get_attribute_value(shape, 'ID') or self._get_attribute_value(shape, 'id')
        
        # Try to get shape type/master
        try:
            master = self._get_attribute_value(shape, 'master_shape')
            if master:
                master_name = self._get_attribute_value(master, 'name')
                shape_info["type"] = str(master_name) if master_name else ""
        except:
            pass
        
        # Extract connections
        shape_info["connections"] = self._extract_connections(shape)
        
        # Extract image if present
        image_data = self._extract_image_from_shape(shape)
        if image_data:
            shape_info["has_image"] = True
            self.logger.info(f"Image found in shape ID: {shape_info['id']}, text: '{shape_info['text']}'")
        
        # Extract sub-shapes recursively (limit depth to avoid infinite recursion)
        if depth < 5:
            try:
                sub_shapes = self._get_shapes_from_object(shape)
                for sub_shape in sub_shapes:
                    sub_info = self._extract_shape_info(sub_shape, depth + 1)
                    # Add ALL sub-shapes, even if they only have images
                    if (sub_info["text"] or sub_info["name"] or sub_info["sub_shapes"] or 
                        sub_info["has_image"] or sub_info["id"]):
                        shape_info["sub_shapes"].append(sub_info)
            except:
                pass
        
        return shape_info

    def _extract_page_data(self, page) -> Dict[str, Any]:
        """
        Extracts data from a Visio page.
        
        Args:
            page: A vsdx.Page object
            
        Returns:
            Dict containing page information
        """
        page_name = self._get_attribute_value(page, 'name', 'Unnamed Page')
        
        page_data = {
            "name": str(page_name) if page_name else "Unnamed Page",
            "shapes": [],
            "connections": [],
            "images_count": 0,
            "page_media": []
        }
        
        # Extract media directly from page (dragged images)
        page_media = self._extract_media_from_page(page)
        if page_media:
            page_data["page_media"] = page_media
            page_data["images_count"] += len(page_media)
            self.logger.info(f"Found {len(page_media)} media items directly on page '{page_name}'")
        
        # Get all shapes from the page
        shapes_to_process = self._get_shapes_from_object(page)
        
        self.logger.info(f"Processing {len(shapes_to_process)} shapes on page '{page_name}'")
        
        # Extract each shape
        for shape in shapes_to_process:
            try:
                shape_info = self._extract_shape_info(shape)
                
                # ALWAYS add the shape to track all shapes, even empty ones
                page_data["shapes"].append(shape_info)
                
                # Collect all connections
                if shape_info["connections"]:
                    page_data["connections"].extend(shape_info["connections"])
                
                # Count images
                if shape_info.get("has_image"):
                    page_data["images_count"] += 1
                    self.logger.info(f"Image in shape found on page '{page_name}'")
                    
            except Exception as e:
                self.logger.warning(f"Failed to extract shape: {e}")
        
        return page_data

    def _sanitize_mermaid_id(self, text: str, shape_id: Any = None) -> str:
        """
        Creates a valid Mermaid node ID from text or shape ID.
        
        Args:
            text: Text to sanitize
            shape_id: Optional shape ID to use if text is empty
            
        Returns:
            Sanitized ID string
        """
        if text:
            # Remove special characters and spaces
            sanitized = ''.join(c if c.isalnum() else '_' for c in text)
            return sanitized[:50]  # Limit length
        elif shape_id:
            return f"shape_{shape_id}"
        return "unknown"

    def _generate_mermaid_diagram(self, page_data: Dict[str, Any]) -> str:
        """
        Generates a Mermaid diagram from page data.
        
        Args:
            page_data: Dictionary containing page information
            
        Returns:
            Mermaid diagram as string
        """
        lines = ["```mermaid", "graph TD"]
        
        # Create a map of shape IDs to their text for labeling
        shape_map = {}
        for shape in page_data["shapes"]:
            if shape.get("id"):
                text = shape.get("text", "")
                if not text and shape.get("has_image"):
                    text = f"Image_{shape['id']}"
                if text:
                    shape_map[str(shape["id"])] = text
            
            # Also map sub-shapes
            for sub_shape in shape.get("sub_shapes", []):
                if sub_shape.get("text"):
                    # Create a unique identifier for sub-shapes
                    sub_id = f"{shape.get('id')}_{sub_shape.get('text')}"
                    shape_map[sub_id] = sub_shape["text"]
        
        # Add page-level images
        if page_data.get('page_media'):
            for idx, media in enumerate(page_data['page_media']):
                node_id = f"page_image_{idx}"
                lines.append(f'    {node_id}["📷 Page Image {idx + 1}"]')
        
        # Add all shapes as nodes
        added_nodes = set()
        for shape in page_data["shapes"]:
            shape_id = shape.get("id")
            text = shape.get("text", "")
            
            # Handle shapes with images but no text
            if not text and shape.get("has_image"):
                text = f"Image_{shape_id}"
            
            if text and shape_id:
                node_id = self._sanitize_mermaid_id(text, shape_id)
                if node_id not in added_nodes:
                    display_text = text.replace('"', "'")
                    # Add image indicator if shape contains an image
                    if shape.get("has_image"):
                        display_text += " 📷"
                    lines.append(f'    {node_id}["{display_text}"]')
                    added_nodes.add(node_id)
        
        # Add connections if any exist
        if page_data.get("connections"):
            lines.append("")
            lines.append("    %% Connections")
            for from_id, to_id in page_data["connections"]:
                if from_id in shape_map and to_id in shape_map:
                    from_node = self._sanitize_mermaid_id(shape_map[from_id], from_id)
                    to_node = self._sanitize_mermaid_id(shape_map[to_id], to_id)
                    lines.append(f"    {from_node} --> {to_node}")
        else:
            # If no explicit connections, try to infer hierarchy from structure
            lines.append("")
            lines.append("    %% Hierarchical structure (inferred)")
            
            # Group shapes by their text patterns (looking for numbered sequences)
            setup_shapes = []
            staging_shapes = []
            finalization_shapes = []
            
            for shape in page_data["shapes"]:
                text = shape.get("text", "").lower()
                if "setup" in text or text.startswith("0") or text.startswith("1"):
                    setup_shapes.append(shape)
                elif "staging" in text or text.startswith("2"):
                    staging_shapes.append(shape)
                elif "finalization" in text or text.startswith("3"):
                    finalization_shapes.append(shape)
            
            # Create flow connections
            if setup_shapes and staging_shapes:
                for s in setup_shapes[:1]:  # Connect first setup to staging
                    for t in staging_shapes[:1]:
                        if s.get("text") and t.get("text"):
                            from_node = self._sanitize_mermaid_id(s["text"], s.get("id"))
                            to_node = self._sanitize_mermaid_id(t["text"], t.get("id"))
                            lines.append(f"    {from_node} --> {to_node}")
            
            if staging_shapes and finalization_shapes:
                for s in staging_shapes[:1]:
                    for t in finalization_shapes[:1]:
                        if s.get("text") and t.get("text"):
                            from_node = self._sanitize_mermaid_id(s["text"], s.get("id"))
                            to_node = self._sanitize_mermaid_id(t["text"], t.get("id"))
                            lines.append(f"    {from_node} --> {to_node}")
        
        lines.append("```")
        return "\n".join(lines)

    def _to_markdown(self, data: Dict[str, Any]) -> str:
        """
        Converts the extracted Visio data to markdown format with Mermaid diagrams.
        
        Args:
            data: The structured data extracted from the Visio file
            
        Returns:
            Markdown formatted string
        """
        markdown = []
        
        # File header
        markdown.append(f"# {data['file_name']}\n")
        
        # Metadata section
        if data.get('metadata') and any(data['metadata'].values()):
            markdown.append("## Metadata\n")
            for key, value in data['metadata'].items():
                if value:
                    markdown.append(f"- **{key.title()}**: {value}")
            markdown.append("")
        
        # Summary
        markdown.append(f"**Total Images Found**: {data.get('total_images', 0)}\n")
        
        # Pages section
        markdown.append(f"## Pages ({len(data['pages'])} total)\n")
        
        for idx, page in enumerate(data['pages'], 1):
            markdown.append(f"### Page {idx}: {page['name']}\n")
            
            if page.get('images_count', 0) > 0:
                markdown.append(f"**Images found on this page**: {page['images_count']}\n")
            
            # Page-level media (dragged images)
            if page.get('page_media'):
                markdown.append(f"#### Page-Level Media ({len(page['page_media'])} items)\n")
                for media_idx, media in enumerate(page['page_media']):
                    markdown.append(f"**Media Item {media_idx + 1}** (Source: {media['source']})")
                    markdown.append("")
            
            # Generate Mermaid diagram for the page
            if page['shapes'] or page.get('page_media'):
                markdown.append("#### Diagram\n")
                mermaid = self._generate_mermaid_diagram(page)
                markdown.append(mermaid)
                markdown.append("")
            
            # Connections summary
            if page.get('connections'):
                markdown.append(f"#### Connections ({len(page['connections'])} total)\n")
                for from_id, to_id in page['connections']:
                    markdown.append(f"- Shape {from_id} → Shape {to_id}")
                markdown.append("")
            
            # Detailed shape information
            if page['shapes']:
                markdown.append(f"#### Detailed Shape Information\n")
                markdown.append(f"**Total shapes**: {len(page['shapes'])}\n")
                
                shape_num = 0
                for shape in page['shapes']:
                    # Show shapes with content OR images
                    if (shape.get('text') or shape.get('name') or shape.get('type') or 
                        shape.get('sub_shapes') or shape.get('has_image')):
                        shape_num += 1
                        markdown.append(f"##### Shape {shape_num}")
                        
                        if shape.get('type'):
                            markdown.append(f"- **Type**: {shape['type']}")
                        
                        if shape.get('name'):
                            markdown.append(f"- **Name**: {shape['name']}")
                        
                        if shape.get('text'):
                            text = shape['text'].strip()
                            markdown.append(f"- **Text**: {text}")
                        
                        if shape.get('id') is not None:
                            markdown.append(f"- **ID**: {shape['id']}")
                        
                        if shape.get('has_image'):
                            markdown.append(f"- **Contains Image**: ✅ Yes")
                        
                        # Handle sub-shapes
                        if shape.get('sub_shapes'):
                            markdown.append(f"- **Sub-shapes**: {len(shape['sub_shapes'])}")
                            for sub_idx, sub_shape in enumerate(shape['sub_shapes'], 1):
                                if sub_shape.get('text') or sub_shape.get('has_image'):
                                    img_marker = " 📷" if sub_shape.get('has_image') else ""
                                    text = sub_shape.get('text', f"Image_{sub_shape.get('id')}")
                                    markdown.append(f"  - {text}{img_marker}")
                        
                        markdown.append("")
                
                if shape_num == 0:
                    markdown.append("*No shapes with content found on this page*\n")
            else:
                markdown.append("*No shapes found on this page*\n")
        
        return "\n".join(markdown)

    def convert(self, file_path: str, output_format: str = "both") -> Union[str, Dict[str, Any], Tuple[str, Dict[str, Any]]]:
        """
        Convert a Visio file to Markdown and/or JSON.
        
        Args:
            file_path: Path to the Visio file (.vsdx)
            output_format: Output format - "markdown", "json", or "both" (default)

        Returns:
            Union[str, Dict, tuple]: 
                - If "markdown": Returns markdown string with Mermaid diagrams
                - If "json": Returns structured JSON dict
                - If "both": Returns tuple of (markdown, json)
        """
        
        try:
            from vsdx import VisioFile
        except ImportError:
            raise ImportError(
                "The 'vsdx' library is required to read Visio files. "
                "Install it with: pip install vsdx"
            )
        
        # Validate file path
        path = Path(file_path)
        if not path.exists():
            raise FileNotFoundError(f"File not found: {file_path}")
        
        if path.suffix.lower() not in ['.vsdx']:
            raise ValueError(
                f"Unsupported file format: {path.suffix}. "
                "Only .vsdx files are supported."
            )
        
        try:
            # Open the Visio file
            self.logger.info(f"Opening Visio file: {file_path}")
            with VisioFile(str(path)) as vis:
                result = {
                    "file_name": path.name,
                    "pages": [],
                    "metadata": {},
                    "total_images": 0
                }
                
                # Extract metadata if available
                try:
                    if hasattr(vis, 'app_properties'):
                        app_props = vis.app_properties
                        result["metadata"] = {
                            "title": getattr(app_props, 'title', ''),
                            "creator": getattr(app_props, 'creator', ''),
                            "company": getattr(app_props, 'company', ''),
                        }
                except:
                    pass
                
                # Extract content from each page
                pages = vis.pages
                if callable(pages):
                    pages = pages()
                
                for page in pages:
                    try:
                        page_data = self._extract_page_data(page)
                        result["pages"].append(page_data)
                        result["total_images"] += page_data.get("images_count", 0)
                    except Exception as e:
                        self.logger.warning(f"Failed to extract page: {e}")
                
                self.logger.info(f"Extracted {len(result['pages'])} pages with {result['total_images']} total images")
                
                # Return based on requested format
                if output_format.lower() == "markdown":
                    return self._to_markdown(result)
                elif output_format.lower() == "json":
                    return result
                else:  # both
                    return (self._to_markdown(result), result)
                
        except Exception as e:
            self.logger.error(f"Failed to read Visio file: {e}")
            raise ValueError(f"Error reading Visio file: {e}")


def main():
    """Main entry point for command-line usage."""
    parser = argparse.ArgumentParser(
        description='Convert Visio (.vsdx) files to Markdown with Mermaid diagrams',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Convert to markdown and display
  python visio_to_markdown_standalone.py diagram.vsdx

  # Convert and save to file
  python visio_to_markdown_standalone.py diagram.vsdx --output diagram.md

  # Convert to JSON
  python visio_to_markdown_standalone.py diagram.vsdx --format json --output diagram.json

  # Verbose output
  python visio_to_markdown_standalone.py diagram.vsdx -v
        """
    )
    
    parser.add_argument(
        'input_file',
        help='Path to the Visio (.vsdx) file'
    )
    
    parser.add_argument(
        '-o', '--output',
        help='Output file path (default: print to stdout)'
    )
    
    parser.add_argument(
        '-f', '--format',
        choices=['markdown', 'json', 'both'],
        default='markdown',
        help='Output format (default: markdown)'
    )
    
    parser.add_argument(
        '-v', '--verbose',
        action='store_true',
        help='Enable verbose logging'
    )
    
    args = parser.parse_args()
    
    # Create converter
    converter = VisioToMarkdownConverter(verbose=args.verbose)
    
    try:
        # Convert the file
        result = converter.convert(args.input_file, output_format=args.format)
        
        # Handle output
        if args.output:
            output_path = Path(args.output)
            
            if args.format == 'both':
                # Save both markdown and json
                md_output, json_output = result
                
                # Save markdown
                md_path = output_path.with_suffix('.md')
                md_path.write_text(md_output, encoding='utf-8')
                print(f"Markdown saved to: {md_path}")
                
                # Save JSON
                json_path = output_path.with_suffix('.json')
                json_path.write_text(json.dumps(json_output, indent=2), encoding='utf-8')
                print(f"JSON saved to: {json_path}")
                
            elif args.format == 'json':
                # Save JSON
                output_path.write_text(json.dumps(result, indent=2), encoding='utf-8')
                print(f"JSON saved to: {output_path}")
                
            else:  # markdown
                # Save markdown
                output_path.write_text(result, encoding='utf-8')
                print(f"Markdown saved to: {output_path}")
        else:
            # Print to stdout
            if args.format == 'both':
                md_output, json_output = result
                print(md_output)
                print("\n" + "="*80 + "\nJSON Output:\n" + "="*80 + "\n")
                print(json.dumps(json_output, indent=2))
            elif args.format == 'json':
                print(json.dumps(result, indent=2))
            else:  # markdown
                print(result)
                
    except Exception as e:
        print(f"Error: {e}", file=__import__('sys').stderr)
        return 1
    
    return 0


if __name__ == '__main__':
    exit(main())
