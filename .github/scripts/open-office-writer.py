#!/usr/bin/env python3
"""
DOCX Assignment Evaluator for the OpenOffice Writer Course at "I.I.S. G. Cena di Ivrea".

This module is designed to evaluate DOCX assignment submissions by comparing student files with a reference solution.
It supports configurable weights and tolerances, and automatically converts ODT files to DOCX using LibreOffice in headless mode.

External Parameters:
  - A JSON configuration file containing weight and tolerance values.
  - A reference subfolder containing the solution files for evaluation.

Author: Francesco Giuseppe Gillio
Date: 25.02.2025
"""

import os
import io
import sys
import csv
import json
import difflib
import zipfile
import subprocess
from PIL import Image
from lxml import etree
from docx import Document


def load_student_registry(registry_path: str) -> dict:
    """
    Load the student registry from a JSON file.

    The registry file should contain student IDs as keys and student names as values.
    
    Args:
        registry_path (str): Path to the JSON file containing the student registry.

    Returns:
        dict: A dictionary with student IDs as keys and student names as values.

    Raises:
        SystemExit: If an error occurs while loading the registry.
    """
    try:
        with open(registry_path, "r") as file:
            registry = json.load(file)
            return registry
    except Exception as e:
        print(f"Error loading student registry: {e}")
        sys.exit(1)


def convert_odt_to_docx(file_path: str) -> str:
    """
    Convert an ODT file to DOCX format using LibreOffice in headless mode.

    If the provided file is not in ODT format, the original file path is returned.
    
    Args:
        file_path (str): The path to the ODT file to be converted.

    Returns:
        str: The path to the converted DOCX file, or the original file path if conversion fails.
    """
    if not file_path.lower().endswith(".odt"):
        return file_path
    docx_file_path = file_path.rsplit('.', 1)[0] + '.docx'
    try:
        subprocess.run(
            [
                'libreoffice',
                '--headless',
                '--convert-to',
                'docx',
                file_path,
                '--outdir',
                os.path.dirname(file_path)
            ],
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL
        )
        return docx_file_path
    except Exception as e:
        print(f"Error converting file {file_path}: {e}")
        return file_path


class DocumentAnalyzer:
    """
    Analyzes the content of DOCX documents, extracting details such as paragraphs, images, tables, and margins.

    This class provides methods to extract and parse different components of a DOCX document:
    - Paragraphs: including text, length, style, formatting (bold, italic, underline), font details, and alignment.
    - Images: including format and dimensions.
    - Tables: including row and column count.
    - Margins: extracted from the underlying XML structure of the document.

    Dependencies:
      - python-docx: for interacting with DOCX files.
      - lxml: for parsing the XML structure.
      - Pillow: for image processing and analysis.
    """

    def __init__(self, file_path: str):
        """
        Initialize the DocumentAnalyzer with a provided DOCX file.

        Args:
            file_path (str): The path to the DOCX file to be analyzed.
        
        Raises:
            ValueError: If the DOCX file cannot be loaded.
        """
        self.docx_path = file_path
        try:
            self.doc = Document(file_path)
        except Exception as e:
            raise ValueError(f"Unable to load DOCX file '{file_path}': {e}")

    def get_paragraph_alignment(self, paragraph) -> str:
        """
        Determine the alignment of a given paragraph.

        Args:
            paragraph: The paragraph object from the DOCX document.

        Returns:
            str: The alignment type (one of "left", "center", "right", "justified", or "unknown").
        """
        alignment_map = {0: "left", 1: "center", 2: "right", 3: "justified"}
        try:
            alignment = paragraph.alignment
            return alignment_map.get(alignment, "unknown")
        except AttributeError:
            return "unknown"

    def get_paragraphs_info(self) -> tuple[list[dict], int]:
        """
        Extract detailed information from each paragraph in the document.

        This includes text, length, style, formatting attributes (bold, italic, underline), font names,
        sizes, and alignment. Additionally, it counts the number of empty paragraphs.

        Returns:
            tuple: A tuple containing:
                - A list of dictionaries with paragraph information.
                - An integer count of empty paragraphs.
        """
        paragraphs_data = []
        empty_lines_count = 0

        for paragraph in self.doc.paragraphs:
            try:
                text = paragraph.text.strip()
            except Exception:
                text = ""
            if not text:
                empty_lines_count += 1
            try:
                paragraph_info = {
                    "text": text,
                    "length": len(text),
                    "style": paragraph.style.name if paragraph.style else "unknown",
                    "bold": any(getattr(run, "bold", False) for run in paragraph.runs),
                    "italic": any(getattr(run, "italic", False) for run in paragraph.runs),
                    "underline": any(getattr(run, "underline", False) for run in paragraph.runs),
                    "font": [run.font.name for run in paragraph.runs if run.font and run.font.name],
                    "size": [run.font.size.pt for run in paragraph.runs if run.font and run.font.size],
                    "alignment": self.get_paragraph_alignment(paragraph)
                }
            except Exception as e:
                paragraph_info = {"text": text, "error": f"Error processing paragraph: {e}"}
            paragraphs_data.append(paragraph_info)

        return paragraphs_data, empty_lines_count

    def get_images_info(self) -> list[dict]:
        """
        Retrieve details about images embedded in the DOCX document.

        This method checks the document's relationships for image parts and uses Pillow to obtain the
        image format and dimensions.

        Returns:
            list: A list of dictionaries, each containing:
                - "format": Image format (e.g., JPEG, PNG).
                - "dimensions": Image dimensions (width, height).
        """
        images_data = []
        for rel in self.doc.part.rels:
            if "image" in self.doc.part.rels[rel].target_ref:
                try:
                    image_part = self.doc.part.rels[rel].target_part
                    image_stream = io.BytesIO(image_part.blob)
                    with Image.open(image_stream) as img:
                        images_data.append({
                            "format": img.format,
                            "dimensions": img.size
                        })
                except Exception as e:
                    images_data.append({
                        "error": f"Error processing image in relation {rel}: {e}"
                    })
        return images_data

    def get_tables_info(self) -> list[dict]:
        """
        Extract table information from the DOCX document.

        This method calculates the number of rows and columns in each table.

        Returns:
            list: A list of dictionaries containing the number of rows and columns for each table.
        """
        tables_data = []
        for table in self.doc.tables:
            try:
                num_rows = len(table.rows)
                num_columns = len(table.columns)
                tables_data.append({
                    "rows": num_rows,
                    "columns": num_columns
                })
            except Exception as e:
                tables_data.append({
                    "error": f"Error processing table: {e}"
                })
        return tables_data

    def get_margins(self) -> dict:
        """
        Extract margin configurations from the DOCX document using its XML structure.

        This method opens the DOCX file as a ZIP archive and parses the "word/document.xml" file
        to extract the margin settings from the "pgMar" element.

        Returns:
            dict: A dictionary containing margin attributes (top, bottom, left, right), or an empty
                  dictionary if margins cannot be extracted.
        """
        margins_data = {}
        try:
            with zipfile.ZipFile(self.docx_path, "r") as docx_zip:
                with docx_zip.open("word/document.xml") as xml_file:
                    xml_tree = etree.parse(xml_file)
                    namespaces = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
                    page_margins = xml_tree.xpath("//w:sectPr/w:pgMar", namespaces=namespaces)
                    if page_margins:
                        margins_data = {attr: page_margins[0].get(attr) for attr in page_margins[0].keys()}
        except Exception as e:
            margins_data = {"error": f"Error extracting margins: {e}"}
        return margins_data


class DocumentComparer:
    """
    Compares two DOCX documents based on their structure and formatting.

    This class uses the DocumentAnalyzer to extract features from both a reference document
    and a student submission, comparing components such as paragraphs, images, tables, and margins.
    A matching percentage and a detailed report are generated for each section.

    The comparison includes:
      - Paragraph text and formatting
      - Image formats and dimensions
      - Table row and column counts
      - Margin settings
    """

    def __init__(self, reference_file: str, test_file: str, config: dict = None):
        """
        Initialize the DocumentComparer with a reference and a student submission.

        Args:
            reference_file (str): Path to the reference DOCX file.
            test_file (str): Path to the student submission DOCX file.
            config (dict, optional): Configuration dictionary containing weights and tolerances.
        """
        self.reference_analyzer = DocumentAnalyzer(reference_file)
        self.test_analyzer = DocumentAnalyzer(test_file)
        self.config = config or {
            "weights": {
                "paragraphs": 0.35,
                "images": 0.25,
                "tables": 0.15,
                "margins": 0.25
            },
            "tolerances": {
                "empty_lines": 0,
                "paragraph_similarity_threshold": 1.00,
                "paragraph_bonus": 0,
                "image_dimension_tolerance": 0,
                "table_rows_tolerance": 0,
                "table_columns_tolerance": 0,
                "margin_tolerance": 0
            }
        }

    def assign_score(self, correct: int, total: int) -> float:
        """
        Compute the percentage score for a section based on the number of correct matches.

        Args:
            correct (int): The number of correct matches.
            total (int): The total number of items compared.

        Returns:
            float: The score percentage (0 to 100).
        """
        return (correct / total) * 100.0 if total else 100.0

    def compare_elements(self, reference_elements: list, test_elements: list, element_name: str) -> tuple[list[str], float]:
        """
        Compare two lists of elements (e.g., images, tables, margins) between documents.

        This method applies tolerance values to compute a match percentage for element types
        like images and tables, or performs exact comparisons for other elements.

        Args:
            reference_elements (list): List of elements from the reference document.
            test_elements (list): List of elements from the student submission.
            element_name (str): The name of the element being compared (e.g., "image", "table", "margins").

        Returns:
            tuple: A tuple containing:
                - A list of strings describing the differences.
                - A match percentage score for the element type.
        """
        differences_list = []
        total_elements = len(reference_elements)

        if total_elements == 0 and len(test_elements) == 0:
            return differences_list, 100.0
        if total_elements == 0 or len(test_elements) == 0:
            return differences_list, 0.0

        sum_score = 0.0
        for index, (ref_elem, test_elem) in enumerate(zip(reference_elements, test_elements)):
            element_score = self.compare_element(ref_elem, test_elem, element_name, index)
            sum_score += element_score

        overall_score = sum_score / total_elements
        return differences_list, overall_score

    def compare_element(self, ref_elem, test_elem, element_name, index) -> float:
        """
        Compare a single element between the reference and student submission.
        
        The logic of this comparison varies depending on the type of element being compared.

        Args:
            ref_elem: The element from the reference document.
            test_elem: The element from the student submission.
            element_name (str): The type of element (e.g., "image", "table", "margins").
            index (int): The index of the element in the list.

        Returns:
            float: The match percentage for the element.
        """
        if element_name.lower() == "image":
            return self.compare_image(ref_elem, test_elem, index)
        elif element_name.lower() == "table":
            return self.compare_table(ref_elem, test_elem, index)
        elif element_name.lower() == "margins":
            return self.compare_margins(ref_elem, test_elem)
        else:
            return self.compare_general(ref_elem, test_elem)

    def compare_image(self, ref_elem, test_elem, index) -> float:
        """
        Compare two images (reference vs test) by their format and dimensions.

        Args:
            ref_elem (dict): The reference image data.
            test_elem (dict): The student submission image data.
            index (int): The index of the image.

        Returns:
            float: The match percentage for the image.
        """
        tolerance = self.config.get("tolerances", {}).get("image_dimension_tolerance", 0)
        sub_total = 3  # Compare format, width, and height
        sub_matches = 0

        if ref_elem.get("format") == test_elem.get("format"):
            sub_matches += 1
        
        ref_dim = ref_elem.get("dimensions")
        test_dim = test_elem.get("dimensions")
        if ref_dim and test_dim and len(ref_dim) == 2 and len(test_dim) == 2:
            if abs(ref_dim[0] - test_dim[0]) <= tolerance:
                sub_matches += 1
            if abs(ref_dim[1] - test_dim[1]) <= tolerance:
                sub_matches += 1

        return (sub_matches / sub_total) * 100

    def compare_table(self, ref_elem, test_elem, index) -> float:
        """
        Compare two tables by their row and column counts.

        Args:
            ref_elem (dict): The reference table data.
            test_elem (dict): The student submission table data.
            index (int): The index of the table.

        Returns:
            float: The match percentage for the table.
        """
        tolerance_rows = self.config.get("tolerances", {}).get("table_rows_tolerance", 0)
        tolerance_cols = self.config.get("tolerances", {}).get("table_columns_tolerance", 0)
        sub_total = 2  # Compare rows and columns
        sub_matches = 0

        if abs(ref_elem.get("rows") - test_elem.get("rows")) <= tolerance_rows:
            sub_matches += 1
        if abs(ref_elem.get("columns") - test_elem.get("columns")) <= tolerance_cols:
            sub_matches += 1

        return (sub_matches / sub_total) * 100

    def compare_margins(self, ref_elem, test_elem) -> float:
        """
        Compare two margin configurations.

        Args:
            ref_elem (dict): The reference margin data.
            test_elem (dict): The student submission margin data.

        Returns:
            float: The match percentage for the margins.
        """
        tolerance = self.config.get("tolerances", {}).get("margin_tolerance", 0)
        sub_keys = list(ref_elem.keys())
        sub_total = len(sub_keys)
        sub_matches = 0

        for key in sub_keys:
            try:
                ref_val = int(ref_elem.get(key))
                test_val = int(test_elem.get(key))
                if abs(ref_val - test_val) <= tolerance:
                    sub_matches += 1
            except ValueError:
                pass

        return (sub_matches / sub_total) * 100 if sub_total > 0 else 100

    def compare_general(self, ref_elem, test_elem) -> float:
        """
        Compare general elements using exact equality.

        Args:
            ref_elem: The reference element.
            test_elem: The student submission element.

        Returns:
            float: 100 if equal, 0 otherwise.
        """
        return 100.0 if ref_elem == test_elem else 0.0

    def generate_markdown_report(self, report_lines: list[str]) -> str:
        """
        Generate a Markdown-formatted report from a list of report lines.

        Args:
            report_lines (list[str]): List of report lines.

        Returns:
            str: The generated Markdown report.
        """
        markdown_report = "# Evaluation Report\n\n"
        
        for line in report_lines:
            if line.startswith("Paragraphs:"):
                markdown_report += f"## Paragraphs\n**Score:** {line.split(':', 1)[1].strip()}\n\n"
            elif line.startswith("Images:"):
                markdown_report += f"## Images\n**Score:** {line.split(':', 1)[1].strip()}\n\n"
            elif line.startswith("Tables:"):
                markdown_report += f"## Tables\n**Score:** {line.split(':', 1)[1].strip()}\n\n"
            elif line.startswith("Margins:"):
                markdown_report += f"## Margins\n**Score:** {line.split(':', 1)[1].strip()}\n\n"
            elif line.startswith("Final Score:"):
                markdown_report += f"## Final Score\n**{line}**\n\n"
            elif line.startswith("- **Paragraph"):
                markdown_report += f"{line}\n"
            elif line.startswith("- **Image") or line.startswith("- **Table") or line.startswith("- **Margin"):
                markdown_report += f"{line}\n"
            elif "additional paragraph" in line:
                markdown_report += f"- **{line.strip()}**\n"
            elif line.startswith("  - **Differences:") or line.startswith("    - **"):
                markdown_report += f"{line}\n"
            else:
                markdown_report += f"- {line}\n"

        return markdown_report

    def compare_documents(self) -> tuple[str, float]:
        """
        Compare the reference and test DOCX documents section by section.

        This method evaluates:
          - Paragraphs: by assessing text and formatting similarity.
          - Images, Tables, and Margins: by directly comparing extracted elements.

        The final score is calculated based on configured weights, and a detailed Markdown report is generated.

        Returns:
            tuple: The Markdown report (str) and the final overall score (float).
        """
        report_lines = []

        ref_paragraphs, _ = self.reference_analyzer.get_paragraphs_info()
        test_paragraphs, _ = self.test_analyzer.get_paragraphs_info()
        para_differences, paragraph_score = self.compare_paragraphs(ref_paragraphs, test_paragraphs)
        report_lines.append(f"Paragraphs: {paragraph_score:.1f}% match")
        report_lines.extend(para_differences)

        image_differences, image_score = self.compare_elements(
            self.reference_analyzer.get_images_info(),
            self.test_analyzer.get_images_info(),
            "image"
        )
        report_lines.append(f"Images: {image_score:.1f}% match")
        report_lines.extend(image_differences)

        table_differences, table_score = self.compare_elements(
            self.reference_analyzer.get_tables_info(),
            self.test_analyzer.get_tables_info(),
            "table"
        )
        report_lines.append(f"Tables: {table_score:.1f}% match")
        report_lines.extend(table_differences)

        margin_differences, margin_score = self.compare_elements(
            [self.reference_analyzer.get_margins()],
            [self.test_analyzer.get_margins()],
            "margins"
        )
        report_lines.append(f"Margins: {margin_score:.1f}% match")
        report_lines.extend(margin_differences)

        weights = self.config.get("weights", {})
        w_paragraphs = weights.get("paragraphs", 0.25)
        w_images = weights.get("images", 0.25)
        w_tables = weights.get("tables", 0.25)
        w_margins = weights.get("margins", 0.25)
        final_score = (
            w_paragraphs * paragraph_score +
            w_images * image_score +
            w_tables * table_score +
            w_margins * margin_score
        )
        final_score = min(final_score, 100)
        report_lines.append(f"\nFinal Score: {final_score:.1f}%")

        return self.generate_markdown_report(report_lines), final_score


class DocxAssignmentEvaluator:
    """
    Evaluates student DOCX assignments based on a reference solution.

    This class compares each student's submission against the reference solution using the DocumentComparer
    and generates individual Markdown reports for each student. It also compiles an overall CSV report.
    """

    def __init__(self, assignment_identifier: str, config: dict = None, registry_path: str = None):
        """
        Initialize the evaluator for a specific assignment.

        Args:
            assignment_identifier (str): Unique identifier for the assignment.
            config (dict, optional): Configuration dictionary containing weights and tolerances.
        """
        self.assignment_id = assignment_identifier
        self.registry = load_student_registry(registry_path)  # Load student registry
        self.assignments_dir = "assignments"
        self.solutions_dir = "solutions"
        self.evaluations_dir = os.path.join("evaluations", self.assignment_id)
        self.assignment_folder = os.path.join(self.assignments_dir, *self.assignment_id.split("/"))

        solution_odt = os.path.join(self.solutions_dir, *self.assignment_id.split("/"), "solution.odt")
        solution_docx = os.path.join(self.solutions_dir, *self.assignment_id.split("/"), "solution.docx")
        if os.path.exists(solution_odt):
            self.solution_file = convert_odt_to_docx(solution_odt)
        elif os.path.exists(solution_docx):
            self.solution_file = solution_docx
        else:
            print(f"Error: Reference solution file in '{os.path.join(self.solutions_dir, self.assignment_id)}' is not available.")
            sys.exit(1)

        self.report_file = os.path.join(self.evaluations_dir, f"REPORT.csv")
        self.config = config

    def verify_resources(self) -> bool:
        """
        Verify the existence of the necessary files and directories for evaluation.

        Returns:
            bool: True if all required resources exist, otherwise False.
        """
        if not os.path.exists(self.assignment_folder):
            print(f"Error: Assignment folder '{self.assignment_folder}' is not available.")
            return False

        if not os.path.exists(self.solution_file):
            print(f"Error: Reference solution file '{self.solution_file}' is not available.")
            return False

        os.makedirs(self.evaluations_dir, exist_ok=True)
        return True

    def run_evaluation(self) -> None:
        """
        Run the evaluation process for each student submission.

        The method converts ODT files to DOCX, compares each student's submission against the reference solution,
        and generates individual Markdown reports. A final CSV report summarizing all results is also generated.
        """
        if not self.verify_resources():
            return

        evaluation_results = []

        for student_id, student_name in self.registry.items():
            surname = student_name.split()[0].upper()

            matched_file = None
            for file in os.listdir(self.assignment_folder):
                if file.endswith(".docx") or file.endswith(".odt"):
                    file_name = os.path.splitext(os.path.basename(file))[0].upper()
                    if file_name == surname:
                        matched_file = file
                        print(f"Submission for {student_name}: {file}")
                        break
                      
            if matched_file:
                student_file_path = os.path.join(self.assignment_folder, matched_file)
                student_file_path = convert_odt_to_docx(student_file_path)
                try:
                    comparer = DocumentComparer(self.solution_file, student_file_path, config=self.config)
                    report, final_score = comparer.compare_documents()
                except Exception as e:
                    print(f"Error processing file {matched_file}: {e}")
                    continue
            else:
                final_score = 0.0
                report = f"# Report for {student_name}\n\nNo submission, score: {final_score}%\n"
                print(f"No submission for {student_name}.")

            evaluation_results.append({
                "Student": student_name,
                "Score (%)": round(final_score, 2)
            })
          
            individual_report_path = os.path.join(self.evaluations_dir, f"{student_name}.md")
            try:
                with open(individual_report_path, "w") as report_file:
                    report_file.write(report)
            except Exception as e:
                print(f"Error writing report for file {file}: {e}")
            
            print(f"Evaluation for {student_name}: {final_score:.1f}% (report at {individual_report_path})")

        try:
            with open(self.report_file, "w", newline="") as csvfile:
                fieldnames = ["Student", "Score (%)"]
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                writer.writeheader()
                sorted_results = sorted(evaluation_results, key=lambda x: x["Student"].split()[-1])
                for result in sorted_results:
                    writer.writerow(result)
            print(f"Overall Evaluation Report available at: {self.report_file}")
        except Exception as e:
            print(f"Error writing CSV report: {e}")


def main():
    """
    Main function to run the assignment evaluation process.
    
    The first command-line argument specifies the assignment identifier,
    and the second argument specifies the path to the student registry file.
    """
    if len(sys.argv) < 3:
        print("Error: Please provide the assignment identifier and the registry file path as input parameters.")
        return

    assignment_identifier = sys.argv[1]
    registry_path = sys.argv[2]
    config = None
  
    if len(sys.argv) >= 4:
        config_file = sys.argv[3]
        try:
            with open(config_file, "r") as cf:
                config = json.load(cf)
                print(f"Loaded configuration: {config}")
        except Exception as e:
            print(f"Error reading configuration file '{config_file}': {e}")
            return

    evaluator = DocxAssignmentEvaluator(assignment_identifier, config=config, registry_path=registry_path)
    evaluator.run_evaluation()


if __name__ == "__main__":
    main()
