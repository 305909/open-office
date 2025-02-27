#!/usr/bin/env python3
"""
CSV/ODS Assignment Evaluator for the OpenOffice Calc Course at "I.I.S. G. Cena di Ivrea".

This module evaluates student submissions for an assignment by comparing each
student's CSV file against a reference solution and an assignment template.
It also supports converting ODS files to CSV format using LibreOffice in headless mode.

The evaluation procedure includes:
    1. Verifying that the required directories and files exist.
    2. Converting ODS files (including solution and assignment template) to CSV format.
    3. Reading the solution and assignment template CSV files.
    4. Performing a cell-by-cell evaluation on cells where the solution differs
       from the assignment template.
    5. Computing a percentage score based on the number of correct responses.
    6. Generating and saving detailed Markdown reports for each student and a
       consolidated CSV report summarizing the evaluation scores for all submissions.

Author: Francesco Giuseppe Gillio
Date: 25.02.2025
"""

import os
import sys
import json
import pandas as pd
import subprocess


def load_student_registry(registry_file_path: str) -> dict:
    """
    Load the student registry from a JSON file.

    The registry file is expected to be a JSON format with the following structure:
    {
        "student_id": "Student Name"
    }

    Args:
        registry_file_path (str): Path to the JSON file containing the student registry.

    Returns:
        dict: A dictionary with student IDs as keys and student names as values.
    """
    try:
        with open(registry_file_path, "r") as file:
            return json.load(file)
    except Exception as e:
        print(f"Error loading student registry: {e}")
        sys.exit(1)


def convert_ods_to_csv(ods_file_path: str) -> str:
    """
    Convert an ODS file to CSV format using LibreOffice in headless mode.

    If the input file is in ODS format, this function will invoke LibreOffice to convert it
    into CSV format. The path to the resulting CSV file is returned. If the input file
    is not in ODS format or the conversion fails, the original file path is returned.

    Args:
        ods_file_path (str): The path to the input file, which may be in ODS format.

    Returns:
        str: The path to the converted CSV file, or the original file path if conversion fails.
    """
    if not ods_file_path.lower().endswith(".ods"):
        return ods_file_path

    csv_file_path = ods_file_path.rsplit('.', 1)[0] + '.csv'

    try:
        subprocess.run(
            [
                'libreoffice', '--headless', '--convert-to', 'csv',
                ods_file_path, '--outdir', os.path.dirname(ods_file_path)
            ],
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL
        )
        return csv_file_path
    except Exception as e:
        print(f"Error converting {ods_file_path}: {e}")
        return ods_file_path


class CsvFileHandler:
    """
    Provides utility functions for handling CSV file operations.

    This class includes methods for reading CSV files into pandas DataFrames
    where all values are treated as strings. This ensures consistency during
    the evaluation process by avoiding issues with mixed data types.
    """

    @staticmethod
    def read_csv_file(csv_file_path: str) -> pd.DataFrame:
        """
        Read a CSV file into a pandas DataFrame, treating all values as strings.

        The CSV file is loaded without a header row, and any missing values
        are replaced with empty strings to maintain consistency during comparisons.

        Args:
            csv_file_path (str): The path to the CSV file to be read.

        Returns:
            pd.DataFrame: A DataFrame containing the CSV data with all columns as strings.
        """
        return pd.read_csv(csv_file_path, header=None, dtype=str).fillna("")


class StudentEvaluator:
    """
    Evaluates an individual student's CSV submission against a reference solution and
    an assignment template.

    The evaluation is performed on a cell-by-cell basis over the shared data range
    between the solution and the student's submission. Only cells where the solution
    value differs from the assignment template are evaluated. For each evaluated cell,
    the student's response is compared against the expected solution, and detailed error
    information is recorded.

    The evaluation results in:
        - A comprehensive Markdown report containing an overview and detailed error information.
    """

    @staticmethod
    def evaluate_submission(
        student_file_path: str,
        student_name: str,
        solution_data: pd.DataFrame,
        assignment_data: pd.DataFrame
    ) -> tuple[float, str]:
        """
        Evaluate a student's CSV submission by comparing it to the solution data on a
        cell-by-cell basis.

        Args:
            student_file_path (str): The path to the student's CSV file.
            student_name (str): The student's name.
            solution_data (pd.DataFrame): A DataFrame containing the correct solution values.
            assignment_data (pd.DataFrame): A DataFrame containing the assignment template values.

        Returns:
            tuple:
                - The computed score percentage (rounded to two decimal places).
                - A detailed Markdown report documenting the cell-by-cell comparisons.
        """
        student_data = CsvFileHandler.read_csv_file(student_file_path)

        score = 0
        total_cells_evaluated = 0
        errors = []

        # Determine the overlapping area for evaluation based on the smallest dimensions
        rows_to_evaluate = min(len(solution_data), len(student_data))
        cols_to_evaluate = min(len(solution_data.columns), len(student_data.columns))

        # Loop through each cell within the overlapping region
        for row_idx in range(rows_to_evaluate):
            for col_idx in range(cols_to_evaluate):
                solution_value = solution_data.iat[row_idx, col_idx]
                assignment_value = assignment_data.iat[row_idx, col_idx]
                student_value = student_data.iat[row_idx, col_idx]

                # Evaluate only if the solution differs from the assignment template
                if solution_value != assignment_value:
                    total_cells_evaluated += 1
                    if solution_value == student_value:
                        score += 1
                    else:
                        errors.append(f"- **Cell ({row_idx + 1}, {col_idx + 1}) mismatch:**")
                        errors.append(f"  - **Expected:** `{solution_value}`")
                        errors.append(f"  - **Student Submission:** `{student_value}`")

        # Calculate the percentage score
        score_percentage = (score / total_cells_evaluated) * 100 if total_cells_evaluated > 0 else 0

        # Construct the detailed Markdown report
        report_lines = []
        report_lines.append(f"# Evaluation Report for {student_name}\n")
        report_lines.append("## Overview\n")
        report_lines.append(f"- **Total Cells Evaluated:** {total_cells_evaluated}")
        report_lines.append(f"- **Correct Answers:** {score}")
        report_lines.append(f"- **Final Score:** {round(score_percentage, 2)}%\n")
        report_lines.append("## Errors\n")
        if errors:
            report_lines.extend(errors)
        else:
            report_lines.append("- No errors found.")

        detailed_report = "\n".join(report_lines)
        return round(score_percentage, 2), detailed_report


class AssignmentEvaluator:
    """
    Manages the evaluation process for student CSV/ODS submissions for a given assignment.

    This class is responsible for:
        - Verifying the existence of necessary directories and files.
        - Evaluating all student submissions (in CSV or ODS format) in the assignment folder.
        - Generating individual detailed Markdown reports for each student.
        - Creating a consolidated CSV report summarizing the evaluation results.
    """

    def __init__(self, assignment_id: str, registry_file_path: str):
        """
        Initialize the AssignmentEvaluator using the provided assignment identifier.

        Args:
            assignment_id (str): The assignment identifier used to locate the corresponding
                                  assignment folder and solution files.
            registry_file_path (str): Path to the JSON file containing the student registry.
        """
        self.assignment_directory = "assignments"
        self.solutions_directory = "solutions"
        self.evaluations_directory = "evaluations"
        self.assignment_id = assignment_id
        self.student_registry = load_student_registry(registry_file_path)

        if not self.assignment_id:
            print("Error: Assignment identifier is mandatory.")
            sys.exit(1)

        self.assignment_folder = os.path.join(self.assignment_directory, *self.assignment_id.split("/"))
        self.evaluation_subfolder = os.path.join(self.evaluations_directory, self.assignment_id)
        os.makedirs(self.evaluation_subfolder, exist_ok=True)

        self.solution_file = self._get_solution_file()
        self.assignment_file = self._get_assignment_template_file()

        self.report_file_path = os.path.join(self.evaluation_subfolder, "REPORT.csv")

    def _get_solution_file(self) -> str:
        """Determine and return the solution file path (preferring .ods over .csv)."""
        solution_path_ods = os.path.join(self.solutions_directory, *self.assignment_id.split("/"), "solution.ods")
        solution_path_csv = os.path.join(self.solutions_directory, *self.assignment_id.split("/"), "solution.csv")

        if os.path.exists(solution_path_ods):
            return convert_ods_to_csv(solution_path_ods)
        elif os.path.exists(solution_path_csv):
            return solution_path_csv
        else:
            print(f"Error: Solution file not available for '{self.assignment_id}'.")
            sys.exit(1)

    def _get_assignment_template_file(self) -> str:
        """Determine and return the assignment template file path (preferring .ods over .csv)."""
        assignment_path_ods = os.path.join(self.solutions_directory, *self.assignment_id.split("/"), "assignment.ods")
        assignment_path_csv = os.path.join(self.solutions_directory, *self.assignment_id.split("/"), "assignment.csv")

        if os.path.exists(assignment_path_ods):
            return convert_ods_to_csv(assignment_path_ods)
        elif os.path.exists(assignment_path_csv):
            return assignment_path_csv
        else:
            print(f"Error: Assignment template file not available for '{self.assignment_id}'.")
            sys.exit(1)

    def verify_resources(self) -> bool:
        """
        Verify the existence of necessary resources for the evaluation process.

        Returns:
            bool: True if all required resources are available, otherwise False.
        """
        if not os.path.exists(self.assignment_folder):
            print(f"Error: Folder '{self.assignment_folder}' not found.")
            return False

        if not os.path.exists(self.solution_file):
            print(f"Error: Solution file '{self.solution_file}' not found.")
            return False

        if not os.path.exists(self.assignment_file):
            print(f"Error: Assignment template file '{self.assignment_file}' not found.")
            return False

        return True

    def run_evaluation(self) -> None:
        """
        Execute the complete evaluation process for all student submissions.

        This includes:
            - Verifying the required directories and files.
            - Evaluating all student submissions.
            - Generating detailed reports for each student.
            - Creating a consolidated CSV report summarizing the results.
        """
        if not self.verify_resources():
            return

        solution_data = CsvFileHandler.read_csv_file(self.solution_file)
        assignment_data = CsvFileHandler.read_csv_file(self.assignment_file)

        evaluation_results = []

        for student_id, student_name in self.student_registry.items():
            student_file = self._find_student_submission(student_name)

            if student_file:
                student_file_path = os.path.join(self.assignment_folder, student_file)
                if student_file_path.lower().endswith(".ods"):
                    student_file_path = convert_ods_to_csv(student_file_path)

                try:
                    score, report = StudentEvaluator.evaluate_submission(
                        student_file_path, student_name, solution_data, assignment_data
                    )
                    print(f"Evaluation for {student_name}: {score}%")
                except Exception as e:
                    print(f"Error processing file {student_file}: {e}")
                    continue

                evaluation_results.append({
                    "Student": student_name,
                    "Score (%)": score
                })

                # Save the individual Markdown report
                self._save_individual_report(student_name, report)
            else:
                evaluation_results.append({
                    "Student": student_name,
                    "Score (%)": 0.0
                })
                self._save_individual_report(student_name, f"# Report for {student_name}\n\nNo submission, score: 0%\n")

        # Create the consolidated CSV report
        self._create_consolidated_report(evaluation_results)

    def _find_student_submission(self, student_name: str) -> str:
        """
        Find the file matching the student's name in the assignment folder.

        Args:
            student_name (str): The student's name.

        Returns:
            str: The matched student file name or None if no file is found.
        """
        surname = student_name.split()[0].upper()
        for file in os.listdir(self.assignment_folder):
            if file.endswith(".csv") or file.endswith(".ods"):
                if os.path.splitext(os.path.basename(file))[0].upper() == surname:
                    return file
        return None

    def _save_individual_report(self, student_name: str, report: str) -> None:
        """Save the individual Markdown report for a student."""
        individual_report_path = os.path.join(self.evaluation_subfolder, f"{student_name}.md")
        try:
            with open(individual_report_path, "w") as report_file:
                report_file.write(report)
        except Exception as e:
            print(f"Error writing report for {student_name}: {e}")

    def _create_consolidated_report(self, evaluation_results: list) -> None:
        """Create and save a consolidated CSV report summarizing the evaluation results."""
        report_df = pd.DataFrame(evaluation_results)
        report_df.to_csv(self.report_file_path, index=False)
        print(f"Overall Evaluation Report saved at: {self.report_file_path}")


def main():
    """
    Main function to execute the evaluation of student submissions for a given assignment.

    The assignment identifier and student registry file path are obtained from the command-line arguments.
    """
    if len(sys.argv) < 3:
        print("Error: Please provide the assignment identifier and the registry file path.")
        return

    assignment_id = sys.argv[1]
    registry_file_path = sys.argv[2]

    evaluator = AssignmentEvaluator(assignment_id, registry_file_path)
    evaluator.run_evaluation()


if __name__ == "__main__":
    main()
