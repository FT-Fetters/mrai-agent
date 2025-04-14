import docx
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
# from docx.enum.section import WD_ORIENTATION, WD_SECTION # Not used
from docx.oxml.ns import qn
# from docx.oxml import OxmlElement # Not used directly now
# import docx.opc.constants # Not used
from docx.shared import Twips # <<< CORRECTED IMPORT for width conversion
from loguru import logger
import os # <<< ADDED IMPORT

# <<< ADDED IMPORTs for table/cell formatting
from docx.enum.table import WD_ALIGN_VERTICAL, WD_ROW_HEIGHT_RULE
# Removed redundant: from docx.oxml.shared import qn (Ensuring this is removed)

from mrai.agent.schema import Tool

# Configure logging (No longer needed with loguru defaults)
# logging.basicConfig(level=logging.INFO)
# logger = logging.getLogger(__name__)

def word_tool_list():
    return [
        ReadWordTool()
    ]

class ReadWordTool(Tool):
    """A tool to read content and basic formatting from Word documents."""

    def __init__(self):
        super().__init__(
            name="read_word",
            description="Reads content and basic formatting from a Word document, including paragraphs, text formatting, and tables.",
            parameters={
                "file_path": Tool.ToolParameter(
                    name="file_path",
                    type="string",
                    description="The path to the Word document file.",
                    required=True
                )
            }
        )

    def _format_value(self, value, unit="", precision=1):
        """Safely formats a numeric value, handling None."""
        if value is None:
            return "Default"
        # Ensure value is numeric before formatting
        try:
            # Attempt conversion to float for formatting
            float_value = float(value)
            return f"{float_value:.{precision}f}{unit}"
        except (ValueError, TypeError):
            # If conversion fails, return the original value as string
            return str(value)

    def _get_run_formatting(self, run):
        """Extracts formatting information for a Run."""
        font = run.font
        formatting = {
            "text": run.text,
            "bold": run.bold,
            "italic": run.italic,
            "underline": run.underline,
            "font_name": font.name,
            # Handle potential AttributeError if size is None
            "font_size_pt": font.size.pt if font.size and hasattr(font.size, 'pt') else None,
            # Handle potential AttributeError if color or rgb is None
            "font_color_rgb": font.color.rgb if font.color and hasattr(font.color, 'rgb') else None,
        }
        return formatting

    def _get_paragraph_formatting(self, paragraph):
        """Extracts formatting information for a Paragraph."""
        para_format = paragraph.paragraph_format
        # Helper to safely access attributes that might be None
        def safe_get_attr(obj, attr, unit):
            val = getattr(obj, attr, None)
            return val.cm if val and hasattr(val, 'cm') and unit == 'cm' else (val.pt if val and hasattr(val, 'pt') and unit == 'pt' else val)

        formatting = {
            "alignment": para_format.alignment,
            "line_spacing": para_format.line_spacing,
            "line_spacing_rule": para_format.line_spacing_rule,
            "first_line_indent_cm": safe_get_attr(para_format, 'first_line_indent', 'cm'),
            "left_indent_cm": safe_get_attr(para_format, 'left_indent', 'cm'),
            "right_indent_cm": safe_get_attr(para_format, 'right_indent', 'cm'),
            "space_before_pt": safe_get_attr(para_format, 'space_before', 'pt'),
            "space_after_pt": safe_get_attr(para_format, 'space_after', 'pt'),
        }
        return formatting

    def _process_paragraph(self, para, index):
        """Processes a single paragraph and returns its formatted string."""
        if not para.text.strip(): # Skip empty paragraphs
            return None

        output = []
        output.append(f"\n### Paragraph {index}")
        output.append(f"**Text:** {para.text}")

        para_format = self._get_paragraph_formatting(para)
        align_str = str(para_format['alignment']) if para_format['alignment'] else 'Default'
        spacing_str = self._format_value(para_format['line_spacing'], "", 1)
        spacing_rule_str = str(para_format['line_spacing_rule']) if para_format['line_spacing_rule'] else 'Default'
        first_indent_str = self._format_value(para_format['first_line_indent_cm'], "cm", 2)
        left_indent_str = self._format_value(para_format['left_indent_cm'], "cm", 2)
        right_indent_str = self._format_value(para_format['right_indent_cm'], "cm", 2)
        space_before_str = self._format_value(para_format['space_before_pt'], "pt", 1)
        space_after_str = self._format_value(para_format['space_after_pt'], "pt", 1)

        output.append(f"**Paragraph Formatting:** Alignment={align_str}, Line Spacing={spacing_str} ({spacing_rule_str}), "
                      f"First Line Indent={first_indent_str}, Left Indent={left_indent_str}, Right Indent={right_indent_str}, "
                      f"Space Before={space_before_str}, Space After={space_after_str}")

        if para.runs:
            output.append("**Run Formatting:**")
            for run in para.runs:
                if not run.text.strip(): # Skip empty runs
                    continue
                run_format = self._get_run_formatting(run)

                rgb_color = run_format['font_color_rgb']
                color_str = 'Default'
                if isinstance(rgb_color, RGBColor): # Check if it's an RGBColor object
                    try:
                        # Format as two-digit hex
                        color_str = f"#{rgb_color[0]:02X}{rgb_color[1]:02X}{rgb_color[2]:02X}"
                    except (TypeError, IndexError, ValueError):
                        logger.warning(f"Could not format RGB color value: {rgb_color}. Using 'Default'.")
                        # Keep color_str as 'Default'
                elif rgb_color is not None: # Handle cases where it might be something else unexpected
                     logger.warning(f"Unexpected color format: {rgb_color}. Using 'Default'.")


                size_str = self._format_value(run_format['font_size_pt'], "pt", 1)
                font_name_str = run_format['font_name'] if run_format['font_name'] else 'Default'
                output.append(f"- `{run_format['text']}`: "
                              f"Font='{font_name_str}', "
                              f"Size={size_str}, "
                              f"Color={color_str}, "
                              f"Bold={'Yes' if run_format['bold'] else 'No'}, "
                              f"Italic={'Yes' if run_format['italic'] else 'No'}, "
                              f"Underline={'Yes' if run_format['underline'] else 'No'}")
        output.append("\n---")
        return "\n".join(output)

    def _process_table(self, table, index):
        """Processes a single table and returns its formatted Markdown string."""
        output = []
        output.append(f"\n### Table {index}")

        # Robustly get column count
        num_cols = 0
        if hasattr(table, 'columns') and table.columns:
             num_cols = len(table.columns)
        elif table.rows and table.rows[0].cells: # Fallback: Use first row's cell count if exists
            num_cols = len(table.rows[0].cells)

        # Try to get table style name
        style_name = "Default"
        if hasattr(table, 'style') and table.style and hasattr(table.style, 'name'):
            style_name = table.style.name
        output.append(f"Table Style: '{style_name}'")

        # <<< ADDED: Check for table-level borders
        try:
            tblPr = table._tbl.tblPr
            if tblPr is not None:
                # Safely find the tblBorders element
                tblBorders_element = tblPr.find(qn('w:tblBorders'))
                if tblBorders_element is not None:
                    # Check for common border types by looking for child elements
                    has_borders = any(
                        tblBorders_element.find(qn(f'w:{tag}')) is not None
                        for tag in ['top', 'bottom', 'left', 'right', 'insideH', 'insideV']
                    )
                    if has_borders:
                        output.append("Table Borders: Defined (specific styles vary)")
        except Exception as e:
            logger.debug(f"Could not check table-level borders for table {index}: {e}")
        # >>> END ADDED

        output.append(f"(Approximate Columns: {num_cols}, Rows: {len(table.rows)})\n")

        formatting_notes = [] # Store formatting notes (row_idx, col_idx_or_0, note)

        if num_cols > 0 and table.rows:
            # <<< ADDED: Process header row height
            self._add_row_formatting_notes(table.rows[0], 1, formatting_notes)
            # >>> END ADDED

            # Build Markdown table header from first row
            header_cells_text = []
            first_row_cells = table.rows[0].cells
            effective_num_cols_header = 0 # Track columns considering spans in header
            col_idx_markdown_header = 0 # Track markdown column index for header
            for i in range(len(first_row_cells)):
                 # Check if we already filled this markdown column due to a previous span
                 if col_idx_markdown_header >= len(header_cells_text):
                     cell = first_row_cells[i]
                     # Extract text
                     cell_text = " ".join(p.text for p in cell.paragraphs).strip().replace("|", "\\|")
                     header_cells_text.append(cell_text)

                     # Check for horizontal span
                     grid_span = 1
                     tcPr = cell._tc.tcPr # Get cell properties element
                     if tcPr is not None:
                         grid_span_elem = tcPr.gridSpan
                         if grid_span_elem is not None and grid_span_elem.val is not None:
                             try:
                                 grid_span = int(grid_span_elem.val)
                             except (ValueError, TypeError):
                                 grid_span = 1

                     # Add formatting notes (use 1-based index for notes)
                     self._add_cell_formatting_notes(cell, 1, col_idx_markdown_header + 1, formatting_notes)
                     if grid_span > 1:
                         formatting_notes.append((1, col_idx_markdown_header + 1, f"Horizontally spans {grid_span} columns"))
                         # Add placeholder cells for Markdown structure
                         header_cells_text.extend([""] * (grid_span - 1))
                         col_idx_markdown_header += grid_span
                     else:
                         col_idx_markdown_header += 1
                 else:
                     # This cell position is covered by a previous span, skip the actual cell
                     # but advance markdown header index if needed (though span handling should cover this)
                     # This logic might need refinement if complex spans exist
                     pass

            # Adjust num_cols based on the effective width found in the header
            num_cols = max(num_cols, len(header_cells_text))

            # Pad header if needed (less likely now but safety)
            while len(header_cells_text) < num_cols:
                 header_cells_text.append(f"Col_{len(header_cells_text)+1}")

            header = "| " + " | ".join(header_cells_text) + " |"
            separator = "|-" + "-|".join(['-' * max(3, len(str(txt))) for txt in header_cells_text]) + "-|"

            output.append(header)
            output.append(separator)

            # Fill table content (iterate through rows *starting from the second row*)
            for row_idx, row in enumerate(table.rows[1:], start=1): # Start from row index 1 (second row)
                # <<< ADDED: Process data row height
                self._add_row_formatting_notes(row, row_idx + 1, formatting_notes)
                # >>> END ADDED

                row_content = []
                cells_to_process = row.cells
                col_idx_actual = 0 # Index for accessing cells_to_process
                col_idx_markdown = 0 # Index for markdown columns, accounting for spans

                while col_idx_markdown < num_cols and col_idx_actual < len(cells_to_process):
                    # Ensure the target markdown column isn't already filled by a previous span
                    if col_idx_markdown >= len(row_content):
                        cell = cells_to_process[col_idx_actual]
                        cell_text = "\n".join([p.text for p in cell.paragraphs])
                        escaped_text = cell_text.replace("|", "\\|").strip()
                        escaped_text = escaped_text.replace('\n', '<br>')
                        row_content.append(escaped_text)

                        # Check for horizontal span
                        grid_span = 1
                        tcPr = cell._tc.tcPr
                        if tcPr is not None:
                             grid_span_elem = tcPr.gridSpan
                             if grid_span_elem is not None and grid_span_elem.val is not None:
                                 try:
                                     grid_span = int(grid_span_elem.val)
                                 except (ValueError, TypeError):
                                     grid_span = 1

                        # Add formatting notes (use row_idx+1 because enumerate starts at 1 here, col_idx_markdown+1)
                        self._add_cell_formatting_notes(cell, row_idx + 1, col_idx_markdown + 1, formatting_notes)
                        if grid_span > 1:
                            formatting_notes.append((row_idx + 1, col_idx_markdown + 1, f"Horizontally spans {grid_span} columns"))
                            # Add placeholder cells for Markdown structure
                            row_content.extend([""] * (grid_span - 1))
                            col_idx_markdown += grid_span
                        else:
                            col_idx_markdown += 1
                    else:
                        # This markdown column was filled by a span from a cell earlier in this row.
                        # We still need to process the *next* actual cell.
                        pass # Let the outer loop increment col_idx_markdown implicitly

                    col_idx_actual += 1

                # Pad row if it has fewer cells than num_cols (due to spans or short rows)
                while len(row_content) < num_cols:
                    row_content.append("")
                    # col_idx_markdown += 1 # Already handled by loop structure

                output.append("| " + " | ".join(row_content) + " |")

        elif table.rows: # Has rows but couldn't determine columns reliably
             output.append("*(Table found, but structure unclear - printing raw row content)*")
             for r_idx, row in enumerate(table.rows):
                 row_text = " | ".join([" ".join(p.text for p in cell.paragraphs).strip() for cell in row.cells])
                 output.append(f"Row {r_idx+1}: {row_text}")
        else:
            output.append("*(Empty Table)*")


        # Add formatting notes if any were collected
        if formatting_notes:
             output.append("\n**Formatting Notes:**")
             # Sort notes primarily by row, then by column (0 for row notes first)
             formatting_notes.sort(key=lambda x: (x[0], x[1]))
             for r, c, note in formatting_notes:
                 if c == 0: # Row-specific note
                     output.append(f"*   Row {r}: {note}")
                 else: # Cell-specific note
                      output.append(f"*   Row {r}, Col {c}: {note}")


        output.append("\n---")
        return "\n".join(output)

    def _add_row_formatting_notes(self, row, row_idx_1_based, notes):
        """Helper to extract and add row height formatting notes."""
        try:
            height = row.height
            height_rule = row.height_rule

            rule_str = None
            if height_rule is not None and height_rule != WD_ROW_HEIGHT_RULE.AUTO:
                # Get the rule name from the enum value
                try:
                    rule_str = WD_ROW_HEIGHT_RULE(height_rule).name
                except ValueError:
                    rule_str = str(height_rule) # Fallback to raw value

            height_str = None
            if height is not None:
                 # height is in Twips, convert to points
                 height_pt = height.pt # Direct conversion using .pt attribute
                 height_str = f"{height_pt:.1f}pt"


            # Add note only if there's non-default info
            if height_str or rule_str:
                note_parts = []
                if height_str:
                    note_parts.append(f"Height={height_str}")
                if rule_str:
                    note_parts.append(f"Rule={rule_str}")
                # Use column 0 to signify a row-level note
                notes.append((row_idx_1_based, 0, ", ".join(note_parts)))

        except Exception as e:
            logger.debug(f"Could not get row height/rule for row index {row_idx_1_based}: {e}")

    def _add_cell_formatting_notes(self, cell, row_idx, col_idx, notes):
        """Helper to extract and add cell formatting notes."""
        tcPr = cell._tc.tcPr # Get cell properties element once

        # --- Text Formatting (First Run) ---
        try:
            if cell.paragraphs and cell.paragraphs[0].runs:
                first_run = cell.paragraphs[0].runs[0]
                if first_run.text.strip(): # Only if the run has text
                    run_format = self._get_run_formatting(first_run)
                    # Format details concisely
                    size_str = self._format_value(run_format['font_size_pt'], "pt", 1)
                    font_name_str = run_format['font_name'] if run_format['font_name'] else 'Default'

                    rgb_color = run_format['font_color_rgb']
                    color_str = 'Default'
                    if isinstance(rgb_color, RGBColor):
                        try:
                            color_str = f"#{rgb_color[0]:02X}{rgb_color[1]:02X}{rgb_color[2]:02X}"
                        except (TypeError, IndexError, ValueError): pass # Keep default on error
                    elif rgb_color is not None: pass # Keep default for unexpected

                    format_parts = [f"Font='{font_name_str}'", f"Size={size_str}"]
                    if color_str != 'Default': format_parts.append(f"Color={color_str}")
                    if run_format['bold']: format_parts.append("Bold")
                    if run_format['italic']: format_parts.append("Italic")
                    if run_format['underline']: format_parts.append("Underline")
                    notes.append((row_idx, col_idx, f"Text Format (Start): {', '.join(format_parts)}"))
        except Exception as e:
            logger.debug(f"Could not get cell text formatting for R{row_idx}C{col_idx}: {e}")


        # --- Horizontal Alignment ---
        try:
            if cell.paragraphs:
                first_para_align = cell.paragraphs[0].alignment
                if first_para_align is not None:
                    align_str = str(first_para_align).split('.')[-1]
                    if align_str != 'LEFT': # LEFT is usually default
                        notes.append((row_idx, col_idx, f"Alignment={align_str}"))
        except Exception as e:
             logger.debug(f"Could not get cell paragraph alignment for R{row_idx}C{col_idx}: {e}")


        # --- Vertical Alignment ---
        try:
            # Access vertical alignment via cell properties (tcPr -> vAlign)
            # tcPr = cell._tc.tcPr # Already defined above
            v_align = None
            if tcPr is not None:
                v_align_elem = tcPr.vAlign
                if v_align_elem is not None and v_align_elem.val is not None:
                    v_align = v_align_elem.val # This is usually the enum value like 'center', 'top', etc.

            is_top = False
            if isinstance(v_align, int): # Check if it's the integer enum value
                is_top = (v_align == WD_ALIGN_VERTICAL.TOP)
            elif isinstance(v_align, str): # Check if it's the string value
                is_top = (v_align.lower() == 'top')

            if v_align is not None and not is_top:
                 valign_str = str(v_align).upper() # Get string representation
                 notes.append((row_idx, col_idx, f"Vertical Alignment={valign_str}"))
        except Exception as e:
             logger.debug(f"Could not get cell vertical alignment for R{row_idx}C{col_idx}: {e}")

        # --- Cell Width ---
        try:
             if tcPr is not None:
                 tcW = tcPr.tcW # Get width element
                 if tcW is not None and tcW.w is not None:
                     width_val = tcW.w
                     width_type = tcW.type
                     if width_type == 'dxa': # Twentieths of a point
                         width_pt = width_val / 20.0
                         notes.append((row_idx, col_idx, f"Width={width_pt:.1f}pt"))
                     elif width_type == 'pct': # Percentage
                         notes.append((row_idx, col_idx, f"Width={width_val}%"))
                     elif width_type == 'auto' or width_type is None:
                         notes.append((row_idx, col_idx, "Width=Auto"))
                     else: # Other types?
                          notes.append((row_idx, col_idx, f"Width={width_val} ({width_type})"))
        except Exception as e:
              logger.debug(f"Could not get cell width for R{row_idx}C{col_idx}: {e}")

        # --- Cell Borders ---
        try:
             if tcPr is not None:
                 # Safely find the tcBorders element
                 tcBorders_element = tcPr.find(qn('w:tcBorders'))
                 if tcBorders_element is not None:
                     # Check if any specific border element exists
                     has_specific_borders = any(
                         tcBorders_element.find(qn(f'w:{tag}')) is not None
                         # We just check for existence, value check might be too complex here
                         for tag in ['top', 'bottom', 'left', 'right'] # Others like tl2br? maybe keep simple
                             # Note: insideH/V might not be typical on tcBorders, mostly on tblBorders
                     )
                     if has_specific_borders:
                         notes.append((row_idx, col_idx, "Has Specific Borders"))
        except Exception as e:
              logger.debug(f"Could not check cell borders for R{row_idx}C{col_idx}: {e}")

         # Note: Detecting shading requires deeper XML parsing (tcPr -> shd)

    def execute(self, file_path: str) -> str:
        """
        Reads the content and formatting information from a Word document,
        processing paragraphs and tables in their original order.

        Args:
            file_path (str): The path to the Word document file.

        Returns:
            str: A string containing the document content and formatting information,
                 formatted using Markdown.
        """
        try:
            document = Document(file_path)
            output = [f"# Word Document Content: {file_path}\n"]
            para_idx = 0
            table_idx = 0

            # Iterate through the document body elements to maintain order
            # Parent element access requires understanding the structure; docx simplifies this.
            # We rely on the assumption that document.paragraphs and document.tables
            # hold references to the objects in the order they appear relative to their type.
            # To get absolute order, we iterate through the XML body's children.
            parent_element = document.element.body
            if parent_element is None:
                 logger.error("Could not access document body element.")
                 return "Error: Could not access document body."

            for child in parent_element:
                # Check if the element is a paragraph ('w:p') or a table ('w:tbl')
                if child.tag == qn('w:p'):
                    if para_idx < len(document.paragraphs):
                        para = document.paragraphs[para_idx]
                        para_output = self._process_paragraph(para, para_idx + 1)
                        if para_output: # Only add if not empty/skipped
                             output.append(para_output)
                        para_idx += 1
                    else:
                         # This might happen if there are paragraph elements not captured by document.paragraphs (e.g., in headers/footers if not handled)
                         logger.warning(f"Found a paragraph element (index {para_idx}) beyond the count in document.paragraphs ({len(document.paragraphs)}). Skipping.")


                elif child.tag == qn('w:tbl'):
                    if table_idx < len(document.tables):
                        table = document.tables[table_idx]
                        table_output = self._process_table(table, table_idx + 1)
                        output.append(table_output)
                        table_idx += 1
                    else:
                         logger.warning(f"Found a table element (index {table_idx}) beyond the count in document.tables ({len(document.tables)}). Skipping.")
                # Note: This loop might not capture elements in headers, footers, text boxes etc.
                # It focuses on the main document body's flow.


            # Simple validation check - might log warnings if counts don't match
            if para_idx != len(document.paragraphs):
                logger.warning(f"Processed {para_idx} paragraphs based on body iteration, but document.paragraphs contains {len(document.paragraphs)}.")
            if table_idx != len(document.tables):
                 logger.warning(f"Processed {table_idx} tables based on body iteration, but document.tables contains {len(document.tables)}.")


            return "\n".join(output)

        except FileNotFoundError:
            logger.error(f"Error: File not found - {file_path}")
            return f"Error: File not found - {file_path}"
        except ImportError as e:
             logger.error(f"Import error, likely missing dependency: {e}. Please ensure 'python-docx' and 'loguru' are installed.")
             return f"Import error: {e}. Make sure 'python-docx' and 'loguru' are installed."
        except Exception as e:
            # Log the full traceback for better debugging using loguru
            logger.exception(f"An unexpected error occurred while reading Word file '{file_path}': {e}")
            return f"An unexpected error occurred while reading Word file '{file_path}': {e}"


if __name__ == '__main__':
    test_file_path = '/Users/xianlindeng/Downloads/乐农〔2025〕69号关于开展2024年度农业“两品”认定奖励申报工作的通知.docx'
    # Check if the file exists before trying to open it
    if os.path.exists(test_file_path):
        print(
            ReadWordTool().execute(test_file_path)
        )
    else:
        print(f"Error: Test file not found at '{test_file_path}'")
        print("Please ensure the test file exists or update the path in the script.")