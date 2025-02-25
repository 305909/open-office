#!/usr/bin/env python3
"""
CSV/ODS Assignment Evaluator for the OpenOffice Calc Course at "I.I.S. G. Cena di Ivrea".

This module evaluates student submissions for an assignment by comparing each
student's CSV file against a reference solution and an assignment template.
It supports converting ODS files to CSV format via LibreOffice in headless mode.
For each student submission, the evaluator produces:
    - A detailed Markdown report containing an overview of evaluated cells and
      an itemized list of cell discrepancies.
    - A consolidated CSV report summarizing the evaluation scores for all
      student submissions.

The evaluation procedure includes:
    1. Verifying that the required directories and files exist.
    2. Converting all ODS files (including solution and assignment template) to CSV.
    3. Reading the solution and assignment template CSV files.
    4. Performing a cell-by-cell evaluation on cells where the solution differs
       from the assignment template.
    5. Computing a percentage score based on the number of correct responses.
    6. Generating and saving detailed Markdown reports for each student as well as
       a consolidated CSV report.

Author: Francesco Giuseppe Gillio
Date: 25.02.2025
"""

import os
import sys
import pandas as pd
import subprocess


def convert_ods_to_csv(file_path: str) -> str:
    """
    Convert an ODS file to CSV format using LibreOffice in headless mode.

    If the input file is an ODS file, this function invokes LibreOffice to perform
    the conversion and returns the path to the resulting CSV file. In case the input
    file is not an ODS file or an error occurs during conversion, the original file
    path is returned unchanged.

    Args:
        file_path (str): The path to the input file that may be in ODS format.

    Returns:
        str: The path to the converted CSV file, or the original file path if
             conversion is not applicable or fails.
    """
    if not file_path.lower().endswith(".ods"):
        return file_path

    # Generate the corresponding CSV file path by replacing the extension.
    csv_file_path = file_path.rsplit('.', 1)[0] + '.csv'

    try:
        subprocess.run(
            [
                'libreoffice', '--headless', '--convert-to', 'csv',
                file_path, '--outdir', os.path.dirname(file_path)
            ],
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL
        )
        return csv_file_path
    except Exception as e:
        print(f"Error converting {file_path}: {e}")
        return file_path


class CsvFileHandler:
    """
    Provides utility functions for handling CSV file operations and related tasks.

    Currently, it includes methods for reading CSV files into a pandas DataFrame
    where all values are treated as strings. This is useful for ensuring uniform
    data types during the evaluation process.
    """

    @staticmethod
    def read_csv_file(file_path: str) -> pd.DataFrame:
        """
        Read a CSV file into a pandas DataFrame with all values as strings.

        The CSV is loaded without a header row, and any missing values are replaced
        with empty strings to ensure consistency during comparisons.

        Args:
            file_path (str): The path to the CSV file to be read.

        Returns:
            pd.DataFrame: A DataFrame containing the CSV data with all columns as strings.
        """
        return pd.read_csv(file_path, header=None, dtype=str).fillna("")


class StudentEvaluator:
    """
    Evaluates an individual student's CSV submission against a reference solution and
    an assignment template.

    The evaluation is performed on a cell-by-cell basis over the common data range
    shared by the solution and the student's submission. Only cells where the solution
    value differs from the assignment template are evaluated. For each evaluated cell,
    the student's response is compared against the expected solution, and detailed error
    information is recorded.

    A comprehensive Markdown report is generated containing:
        - An overview section summarizing the total number of cells evaluated,
          the number of correct responses, and the final score.
        - An error section that lists each cell mismatch with the expected and
          submitted values.
    """

    @staticmethod
    def evaluate_student_file(
        student_file: str,
        solution_data: pd.DataFrame,
        assignment_data: pd.DataFrame
    ) -> tuple[str, str, float, str]:
        """
        Evaluate a student's CSV submission by comparing it to the solution data on a
        cell-by-cell basis.

        The evaluation procedure:
            - Extracts the student's first name and surname from the file name.
            - Reads the student's CSV submission.
            - Iterates over each cell within the overlapping dimensions of the solution
              and student's data.
            - Considers only cells where the solution value is different from the assignment template.
            - Increments the score if the student's cell value matches the expected solution.
            - Records detailed error messages for each cell mismatch.
            - Calculates a final percentage score based on the number of correctly answered cells.
            - Generates a Markdown report detailing the evaluation process and results.

        Args:
            student_file (str): The path to the student's CSV file.
            solution_data (pd.DataFrame): A DataFrame containing the correct solution values.
            assignment_data (pd.DataFrame): A DataFrame containing the assignment template values.

        Returns:
            tuple:
                - The student's first name (capitalized).
                - The student's surname (capitalized).
                - The computed score percentage (rounded to two decimal places).
                - A detailed Markdown report (str) documenting the cell-by-cell comparisons.
        """
        # Extract the student's name from the filename assuming "firstname-surname.csv"
        base_name = os.path.splitext(os.path.basename(student_file))[0]
        name_parts = base_name.split('-')
        student_first_name, student_last_name = [part.lower().capitalize() for part in name_parts]

        # Read the student's CSV file into a DataFrame
        student_data = CsvFileHandler.read_csv_file(student_file)

        score = 0
        total_evaluated_cells = 0
        errors = []

        # Determine the overlapping area for evaluation based on the smallest dimensions
        rows_to_compare = min(len(solution_data), len(student_data))
        cols_to_compare = min(len(solution_data.columns), len(student_data.columns))

        # Loop through each cell within the overlapping region
        for i in range(rows_to_compare):
            for j in range(cols_to_compare):
                sol_value = solution_data.iat[i, j]
                assign_value = assignment_data.iat[i, j]
                student_value = student_data.iat[i, j]
                # Evaluate only if the solution cell differs from the assignment template
                if sol_value != assign_value:
                    total_evaluated_cells += 1
                    if sol_value == student_value:
                        score += 1
                    else:
                        errors.append(f"- **Cell ({i + 1}, {j + 1}) mismatch:**")
                        errors.append(f"  - **Result:** `{sol_value}`")
                        errors.append(f"  - **Student Submission:** `{student_value}`")

        # Calculate the percentage score
        percentage = (score / total_evaluated_cells) * 100 if total_evaluated_cells > 0 else 0

        # Build the detailed Markdown report
        report_lines = []
        report_lines.append(
            f"# Evaluation Report for {student_first_name} {student_last_name}\n"
        )
        report_lines.append("## Overview\n")
        report_lines.append(f"- **Total Evaluation Cells:** {total_evaluated_cells}")
        report_lines.append(f"- **Correct Answers:** {score}")
        report_lines.append(f"- **Final Score:** {round(percentage, 2)}%\n")
        report_lines.append("## Errors\n")
        if errors:
            report_lines.extend(errors)
        else:
            report_lines.append("- No errors found.")

        detailed_report = "\n".join(report_lines)
        return student_first_name, student_last_name, round(percentage, 2), detailed_report


class AssignmentEvaluator:
    """
    Manages the evaluation process for student CSV/ODS submissions for a given assignment.

    This class is responsible for:
        - Determining the assignment folder based on a mandatory assignment identifier.
        - Validating the existence of required directories and files.
        - Evaluating all student submissions (in CSV or ODS format) within the assignment folder.
        - Generating a consolidated CSV report summarizing the evaluation scores.
        - Creating an individual, detailed Markdown report for each student.
    """

    def __init__(self, assignment_arg: str = None):
        """
        Initialize the AssignmentEvaluator using the provided assignment identifier.

        The assignment identifier is mandatory and is used to locate the corresponding
        assignment folder as well as to create a dedicated subfolder in the evaluations
        directory for storing report files.

        Args:
            assignment_arg (str): The assignment identifier.
        """
        self.assignments_directory = "assignments"
        self.solutions_directory = "solutions"
        self.evaluations_directory = "evaluations"
        self.assignment = assignment_arg

        if not self.assignment:
            print("Error: Assignment identifier is mandatory.")
            sys.exit(1)

        self.assignment_folder = os.path.join(self.assignments_directory, *self.assignment.split("/"))

        # Create a subfolder in the evaluations directory for the current assignment
        self.evaluation_subfolder = os.path.join(self.evaluations_directory, self.assignment)
        os.makedirs(self.evaluation_subfolder, exist_ok=True)

        # Determine the solution file by preferring solution.ods over solution.csv
        solution_path_ods = os.path.join(
            self.solutions_directory, *self.assignment.split("/"), "solution.ods"
        )
        solution_path_csv = os.path.join(
            self.solutions_directory, *self.assignment.split("/"), "solution.csv"
        )
        if os.path.exists(solution_path_ods):
            self.solution_file = convert_ods_to_csv(solution_path_ods)
        elif os.path.exists(solution_path_csv):
            self.solution_file = solution_path_csv
        else:
            print(
                f"Error: Solution file not available in "
                f"'{os.path.join(self.solutions_directory, self.assignment)}'."
            )
            sys.exit(1)

        # Determine the assignment template file by preferring assignment.ods over assignment.csv
        assignment_path_ods = os.path.join(
            self.solutions_directory, *self.assignment.split("/"), "assignment.ods"
        )
        assignment_path_csv = os.path.join(
            self.solutions_directory, *self.assignment.split("/"), "assignment.csv"
        )
        if os.path.exists(assignment_path_ods):
            self.assignment_file = convert_ods_to_csv(assignment_path_ods)
        elif os.path.exists(assignment_path_csv):
            self.assignment_file = assignment_path_csv
        else:
            print(
                f"Error: Assignment template file not available in "
                f"'{os.path.join(self.solutions_directory, self.assignment)}'."
            )
            sys.exit(1)

        # Define the path for the consolidated CSV report
        self.report_file = os.path.join(self.evaluation_subfolder, f"Report.csv")

    def verify_resources(self) -> bool:
        """
        Verify that all necessary directories and files exist for the evaluation process.

        This includes checking for the assignment folder, the solution CSV file, and the
        assignment template CSV file.

        Returns:
            bool: True if all required resources are available; otherwise, False.
        """
        if not self.assignment:
            print("Error: Assignment identifier is not available.")
            return False

        if not os.path.exists(self.assignment_folder):
            print(f"Error: Folder '{self.assignment_folder}' not available.")
            return False

        if not os.path.exists(self.solution_file):
            print(f"Error: File '{self.solution_file}' not available.")
            return False

        if not os.path.exists(self.assignment_file):
            print(f"Error: File '{self.assignment_file}' not available.")
            return False

        return True

    def run_evaluation(self) -> None:
        """
        Execute the complete evaluation process for all student CSV/ODS submissions.

        The process involves:
            - Verifying that required resources (directories and files) are present.
            - Reading the solution and assignment template CSV files into DataFrames.
            - Iterating over each file in the assignment folder. For each student file:
                - Converting ODS files to CSV if necessary.
                - Evaluating the student's submission via a cell-by-cell comparison.
                - Generating and saving an individual Markdown report.
            - Creating a consolidated CSV report that summarizes the evaluation results.
        """
        if not self.verify_resources():
            return

        solution_data = CsvFileHandler.read_csv_file(self.solution_file)
        assignment_data = CsvFileHandler.read_csv_file(self.assignment_file)

        evaluation_results = []

        # Process each student file found in the assignment folder.
        for file in os.listdir(self.assignment_folder):
            if file.endswith(".csv") or file.endswith(".ods"):
                student_file_path = os.path.join(self.assignment_folder, file)
                if file.lower().endswith(".ods"):
                    student_file_path = convert_ods_to_csv(student_file_path)
                first_name, last_name, score, detailed_report = StudentEvaluator.evaluate_student_file(
                    student_file_path, solution_data, assignment_data
                )
                evaluation_results.append({
                    "Name": first_name,
                    "Surname": last_name,
                    "Score (%)": score
                })
                print(f"Assessment {file}: {score}%")

                individual_report_filename = f"{first_name}-{last_name}.md"
                individual_report_path = os.path.join(self.evaluation_subfolder, individual_report_filename)
                with open(individual_report_path, "w", encoding="utf-8") as md_file:
                    md_file.write(detailed_report)
                print(f"Detailed report generated for {file}: {individual_report_path}")

        # Create a consolidated CSV report with all evaluation results
        report_dataframe = pd.DataFrame(evaluation_results)
        report_dataframe.to_csv(self.report_file, index=False)
        print(f"Consolidated Evaluation Report generated at: {self.report_file}")


def main():
    """
    Main function to execute the evaluation of student CSV/ODS submissions for an assignment.

    The assignment identifier is obtained from the command-line arguments.
    """
    assignment_identifier = sys.argv[1] if len(sys.argv) > 1 else None
  
    evaluator = AssignmentEvaluator(assignment_identifier)
    evaluator.run_evaluation()


if __name__ == "__main__":
    main()
