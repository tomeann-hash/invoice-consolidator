
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
import pandas as pd
import re


# ============================================================
# CONFIGURATION
# ============================================================

REQUIRED_COLUMNS = [
    "Bill to",
    "Invoice Number",
    "Invoice Date",
    "Total",
]


# ============================================================
# COLUMN HELPERS
# ============================================================

def normalize_column(name):
    """
    Normalize a column name for matching.

    Example:
        'Invoice Number' -> 'invoicenumber'
        ' invoice-number ' -> 'invoicenumber'
    """
    return re.sub(
        r"[^a-z0-9]",
        "",
        str(name).strip().lower()
    )


def find_column(df, target):
    """
    Find a DataFrame column that matches the target name.
    """

    target_norm = normalize_column(target)

    for col in df.columns:
        if normalize_column(col) == target_norm:
            return col

    return None


# ============================================================
# BILL TO CLEANING
# ============================================================

def clean_bill_to(value):
    """
    Keep only the first non-empty line from the Bill to value.

    Example:

        ABC Company
        123 Main Street
        Chennai
        India

    Becomes:

        ABC Company
    """

    if pd.isna(value):
        return ""

    text = str(value)

    text = text.replace("\r\n", "\n")
    text = text.replace("\r", "\n")

    lines = [
        line.strip()
        for line in text.split("\n")
        if line.strip()
    ]

    if lines:
        return lines[0]

    return ""


# ============================================================
# EXCEL EXTRACTION
# ============================================================

def extract_from_excel(source_file):
    """
    Extract required data from an Excel file.
    """

    df = pd.read_excel(source_file)

    mapping = {}

    for target in REQUIRED_COLUMNS:

        found = find_column(
            df,
            target
        )

        if not found:
            available_columns = ", ".join(
                map(str, df.columns)
            )

            raise ValueError(
                f"Could not find required column "
                f"'{target}' in:\n{source_file}\n\n"
                f"Available columns:\n"
                f"{available_columns}"
            )

        mapping[target] = found

    result = df[
        [
            mapping[column]
            for column in REQUIRED_COLUMNS
        ]
    ].copy()

    result.columns = REQUIRED_COLUMNS

    result["Bill to"] = (
        result["Bill to"]
        .apply(clean_bill_to)
    )

    # Remove completely empty rows
    result = result[
        result["Bill to"].ne("")
        | result["Invoice Number"].notna()
        | result["Invoice Date"].notna()
        | result["Total"].notna()
    ].copy()

    return result


# ============================================================
# PDF TEXT EXTRACTION
# ============================================================

def extract_pdf_text(pdf_file):
    """
    Extract text from a text-based PDF.

    Requires:
        pip install pypdf
    """

    try:
        from pypdf import PdfReader
    except ImportError:
        raise ImportError(
            "PDF support requires the 'pypdf' package.\n\n"
            "Install it using:\n"
            "pip install pypdf"
        )

    reader = PdfReader(pdf_file)

    pages = []

    for page in reader.pages:

        text = page.extract_text()

        if text:
            pages.append(text)

    return "\n".join(pages)


# ============================================================
# PDF FIELD EXTRACTION HELPERS
# ============================================================

def clean_extracted_value(value):
    """
    Clean extracted PDF field values.
    """

    if not value:
        return ""

    value = value.replace("\r\n", "\n")
    value = value.replace("\r", "\n")

    value = re.sub(
        r"[ \t]+",
        " ",
        value
    )

    return value.strip()


def extract_field_from_text(
    text,
    patterns
):
    """
    Try multiple regular expression patterns
    to find a field value.
    """

    for pattern in patterns:

        match = re.search(
            pattern,
            text,
            flags=re.IGNORECASE
            | re.MULTILINE
        )

        if match:

            value = match.group(1)

            value = clean_extracted_value(
                value
            )

            if value:
                return value

    return ""


def extract_bill_to_from_pdf(text):
    """
    Attempt to extract the Bill to value from a PDF.

    It supports common labels such as:

        Bill to:
        Bill To
        Billed To
        Customer
        Customer Name
    """

    patterns = [

        # Bill to: ABC Company
        r"Bill\s*To\s*[:\-]\s*([^\n]+)",

        # Billed To: ABC Company
        r"Billed\s*To\s*[:\-]\s*([^\n]+)",

        # Customer: ABC Company
        r"Customer\s*[:\-]\s*([^\n]+)",

        # Customer Name: ABC Company
        r"Customer\s*Name\s*[:\-]\s*([^\n]+)",
    ]

    value = extract_field_from_text(
        text,
        patterns
    )

    return clean_bill_to(value)


def extract_invoice_number_from_pdf(text):
    """
    Attempt to extract invoice number.
    """

    patterns = [

        r"Invoice\s*Number\s*[:#\-]?\s*([^\n]+)",

        r"Invoice\s*No\.?\s*[:#\-]?\s*([^\n]+)",

        r"Invoice\s*#\s*[:\-]?\s*([^\n]+)",

        r"Inv(?:oice)?\s*No\.?\s*[:#\-]?\s*([^\n]+)",
    ]

    return extract_field_from_text(
        text,
        patterns
    )


def extract_invoice_date_from_pdf(text):
    """
    Attempt to extract invoice date.
    """

    patterns = [

        r"Invoice\s*Date\s*[:\-]?\s*([^\n]+)",

        r"Date\s*of\s*Invoice\s*[:\-]?\s*([^\n]+)",

        r"Inv(?:oice)?\s*Date\s*[:\-]?\s*([^\n]+)",
    ]

    return extract_field_from_text(
        text,
        patterns
    )


def extract_total_from_pdf(text):
    """
    Attempt to extract invoice total.

    Supports common labels such as:

        Total
        Grand Total
        Invoice Total
        Amount Due
        Total Amount
    """

    patterns = [

        r"Grand\s*Total\s*[:\-]?\s*([^\n]+)",

        r"Invoice\s*Total\s*[:\-]?\s*([^\n]+)",

        r"Total\s*Amount\s*[:\-]?\s*([^\n]+)",

        r"Amount\s*Due\s*[:\-]?\s*([^\n]+)",

        r"Total\s*[:\-]?\s*([^\n]+)",
    ]

    return extract_field_from_text(
        text,
        patterns
    )


# ============================================================
# PDF EXTRACTION
# ============================================================

def extract_from_pdf(source_file):
    """
    Extract required fields from a PDF invoice.
    """

    text = extract_pdf_text(
        source_file
    )

    if not text.strip():
        raise ValueError(
            "No readable text was found in this PDF.\n\n"
            "The PDF may be scanned/image-based and may "
            "require OCR."
        )

    bill_to = extract_bill_to_from_pdf(
        text
    )

    invoice_number = extract_invoice_number_from_pdf(
        text
    )

    invoice_date = extract_invoice_date_from_pdf(
        text
    )

    total = extract_total_from_pdf(
        text
    )

    result = pd.DataFrame(
        [
            {
                "Bill to": bill_to,
                "Invoice Number": invoice_number,
                "Invoice Date": invoice_date,
                "Total": total,
            }
        ]
    )

    # Return the row even if some fields are empty,
    # provided at least some invoice information was found.
    if not any(
        [
            bill_to,
            invoice_number,
            invoice_date,
            total,
        ]
    ):
        raise ValueError(
            "Could not identify invoice fields in this PDF."
        )

    return result


# ============================================================
# PROCESS ONE FILE
# ============================================================

def process_file(source_file):
    """
    Automatically determine whether the selected file
    is Excel or PDF and extract the required data.
    """

    extension = Path(
        source_file
    ).suffix.lower()

    if extension in [
        ".xlsx",
        ".xls",
    ]:

        return extract_from_excel(
            source_file
        )

    elif extension == ".pdf":

        return extract_from_pdf(
            source_file
        )

    else:

        raise ValueError(
            f"Unsupported file type:\n{source_file}"
        )


# ============================================================
# PROCESS MULTIPLE FILES
# ============================================================

def extract_from_multiple_files(
    source_files
):
    """
    Process multiple Excel and PDF files.

    Returns:
        combined DataFrame
        successful file count
        failed file list
    """

    all_data = []

    successful_files = 0

    failed_files = []

    for source_file in source_files:

        try:

            data = process_file(
                source_file
            )

            # Add source filename for tracking
            data.insert(
                0,
                "Source File",
                Path(source_file).name
            )

            all_data.append(
                data
            )

            successful_files += 1

        except Exception as e:

            failed_files.append(
                (
                    Path(source_file).name,
                    str(e)
                )
            )

    if all_data:

        combined = pd.concat(
            all_data,
            ignore_index=True
        )

    else:

        combined = pd.DataFrame(
            columns=[
                "Source File"
            ] + REQUIRED_COLUMNS
        )

    return (
        combined,
        successful_files,
        failed_files
    )


# ============================================================
# CONSOLIDATION
# ============================================================

def consolidate(
    source_files,
    consolidated_file
):
    """
    Process multiple Excel/PDF files and consolidate
    them into one Excel file.

    Duplicate Invoice Numbers are skipped.
    """

    new_data, successful_files, failed_files = (
        extract_from_multiple_files(
            source_files
        )
    )

    # ========================================================
    # READ EXISTING CONSOLIDATED FILE
    # ========================================================

    if Path(
        consolidated_file
    ).exists():

        existing = pd.read_excel(
            consolidated_file
        )

        # Ensure required columns exist
        for column in REQUIRED_COLUMNS:

            if column not in existing.columns:

                existing[column] = ""

        # Add Source File column if missing
        if "Source File" not in existing.columns:

            existing.insert(
                0,
                "Source File",
                ""
            )

        existing = existing[
            [
                "Source File"
            ] + REQUIRED_COLUMNS
        ]

        # Combine existing and new data
        combined = pd.concat(
            [
                existing,
                new_data
            ],
            ignore_index=True
        )

    else:

        combined = new_data

    # ========================================================
    # REMOVE DUPLICATE INVOICE NUMBERS
    # ========================================================

    # Only remove duplicates when Invoice Number
    # is actually available.
    #
    # Blank invoice numbers are NOT treated as duplicates.

    invoice_numbers = (
        combined["Invoice Number"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    has_invoice_number = (
        invoice_numbers != ""
    )

    with_invoice_number = combined[
        has_invoice_number
    ].copy()

    without_invoice_number = combined[
        ~has_invoice_number
    ].copy()

    # Keep the first occurrence
    with_invoice_number = (
        with_invoice_number
        .drop_duplicates(
            subset=[
                "Invoice Number"
            ],
            keep="first"
        )
    )

    # Recombine
    combined = pd.concat(
        [
            with_invoice_number,
            without_invoice_number
        ],
        ignore_index=True
    )

    # ========================================================
    # SAVE FILE
    # ========================================================

    combined.to_excel(
        consolidated_file,
        index=False
    )

    return (
        len(new_data),
        len(combined),
        successful_files,
        failed_files
    )


# ============================================================
# SELECT MULTIPLE SOURCE FILES
# ============================================================

def choose_sources():
    """
    Select multiple Excel and PDF files.
    """

    paths = filedialog.askopenfilenames(
        title=(
            "Select Excel and PDF invoice files"
        ),
        filetypes=[
            (
                "Excel and PDF files",
                "*.xlsx *.xls *.pdf"
            ),
            (
                "Excel files",
                "*.xlsx *.xls"
            ),
            (
                "PDF files",
                "*.pdf"
            ),
            (
                "All files",
                "*.*"
            ),
        ]
    )

    if paths:

        # Save selected paths
        source_files_var.set(
            paths
        )

        # Display file names in the listbox
        source_listbox.delete(
            0,
            tk.END
        )

        for path in paths:

            source_listbox.insert(
                tk.END,
                path
            )


# ============================================================
# CLEAR SOURCE FILES
# ============================================================

def clear_sources():
    """
    Clear all selected source files.
    """

    source_files_var.set(
        ""
    )

    source_listbox.delete(
        0,
        tk.END
    )


# ============================================================
# SELECT EXISTING CONSOLIDATED FILE
# ============================================================

def choose_consolidated():
    """
    Select an existing consolidated Excel file.
    """

    path = filedialog.askopenfilename(
        title=(
            "Select existing consolidated Excel file"
        ),
        filetypes=[
            (
                "Excel files",
                "*.xlsx *.xls"
            )
        ]
    )

    if path:

        consolidated_var.set(
            path
        )


# ============================================================
# CREATE NEW CONSOLIDATED FILE
# ============================================================

def create_new_consolidated():
    """
    Create a new consolidated Excel file.
    """

    path = filedialog.asksaveasfilename(
        title=(
            "Create consolidated Excel file"
        ),
        defaultextension=".xlsx",
        filetypes=[
            (
                "Excel files",
                "*.xlsx"
            )
        ]
    )

    if path:

        consolidated_var.set(
            path
        )


# ============================================================
# RUN CONSOLIDATION
# ============================================================

def run_consolidation():
    """
    Validate inputs and start consolidation.
    """

    source_text = (
        source_files_var.get().strip()
    )

    consolidated = (
        consolidated_var.get().strip()
    )

    # ========================================================
    # CHECK SOURCE FILES
    # ========================================================

    if not source_text:

        messagebox.showwarning(
            "Missing source files",
            (
                "Please select one or more "
                "Excel or PDF files."
            )
        )

        return

    # Convert stored tuple/string representation
    # into a list of files.
    source_files = list(
        selected_source_files
    )

    if not source_files:

        messagebox.showwarning(
            "Missing source files",
            (
                "Please select one or more "
                "Excel or PDF files."
            )
        )

        return

    # ========================================================
    # CHECK CONSOLIDATED FILE
    # ========================================================

    if not consolidated:

        messagebox.showwarning(
            "Missing consolidated file",
            (
                "Please select an existing consolidated "
                "file or create a new one."
            )
        )

        return

    # ========================================================
    # CHECK FILES EXIST
    # ========================================================

    missing_files = [
        file
        for file in source_files
        if not Path(file).exists()
    ]

    if missing_files:

        messagebox.showerror(
            "File Not Found",
            (
                "The following source files "
                "could not be found:\n\n"
                + "\n".join(
                    missing_files
                )
            )
        )

        return

    # ========================================================
    # RUN
    # ========================================================

    try:

        (
            extracted_rows,
            total_rows,
            successful_files,
            failed_files
        ) = consolidate(
            source_files,
            consolidated
        )

        # ====================================================
        # SUCCESS MESSAGE
        # ====================================================

        message = (
            "Extraction and consolidation completed.\n\n"
            f"Files selected: {len(source_files)}\n"
            f"Files processed successfully: {successful_files}\n"
            f"Rows extracted: {extracted_rows}\n"
            f"Total rows in consolidated file: {total_rows}\n\n"
            f"Saved to:\n{consolidated}"
        )

        # ====================================================
        # FAILED FILES
        # ====================================================

        if failed_files:

            message += (
                "\n\nFiles that could not be processed:"
            )

            for filename, error in failed_files:

                message += (
                    f"\n\n{filename}\n"
                    f"Reason: {error}"
                )

        messagebox.showinfo(
            "Consolidation Completed",
            message
        )

    except Exception as e:

        messagebox.showerror(
            "Error",
            str(e)
        )


# ============================================================
# MAIN WINDOW
# ============================================================

root = tk.Tk()

root.title(
    "Invoice Excel & PDF Consolidator"
)

root.geometry(
    "850x600"
)

root.resizable(
    False,
    False
)


# ============================================================
# VARIABLES
# ============================================================

source_files_var = tk.StringVar()

consolidated_var = tk.StringVar()

# Keep actual selected file paths
selected_source_files = []


# ============================================================
# TITLE
# ============================================================

tk.Label(
    root,
    text=(
        "Invoice Excel & PDF Consolidator"
    ),
    font=(
        "Arial",
        18,
        "bold"
    )
).pack(
    pady=(20, 5)
)


# ============================================================
# DESCRIPTION
# ============================================================

tk.Label(
    root,
    text=(
        "Select multiple Excel and PDF invoice files. "
        "The application extracts Bill to, Invoice Number, "
        "Invoice Date and Total, then consolidates them "
        "into one Excel file."
    ),
    wraplength=760
).pack(
    pady=(0, 15)
)


# ============================================================
# SOURCE FILE FRAME
# ============================================================

source_frame = tk.LabelFrame(
    root,
    text="Source Invoice Files",
    padx=10,
    pady=10
)

source_frame.pack(
    fill="both",
    padx=30,
    pady=5
)


# ============================================================
# SOURCE BUTTONS
# ============================================================

button_frame = tk.Frame(
    source_frame
)

button_frame.pack(
    fill="x",
    pady=(0, 8)
)


def choose_sources_and_store():
    """
    Wrapper that stores selected files globally.
    """

    global selected_source_files

    paths = filedialog.askopenfilenames(
        title=(
            "Select Excel and PDF invoice files"
        ),
        filetypes=[
            (
                "Excel and PDF files",
                "*.xlsx *.xls *.pdf"
            ),
            (
                "Excel files",
                "*.xlsx *.xls"
            ),
            (
                "PDF files",
                "*.pdf"
            ),
        ]
    )

    if paths:

        selected_source_files = list(
            paths
        )

        source_files_var.set(
            f"{len(paths)} file(s) selected"
        )

        source_listbox.delete(
            0,
            tk.END
        )

        for path in paths:

            source_listbox.insert(
                tk.END,
                path
            )


tk.Button(
    button_frame,
    text="Select Multiple Excel / PDF Files",
    command=choose_sources_and_store,
    width=35
).pack(
    side="left",
    padx=5
)


tk.Button(
    button_frame,
    text="Clear",
    command=clear_sources,
    width=12
).pack(
    side="left",
    padx=5
)


# ============================================================
# SELECTED FILE COUNT
# ============================================================

tk.Label(
    source_frame,
    textvariable=source_files_var,
    anchor="w"
).pack(
    fill="x",
    pady=(0, 5)
)


# ============================================================
# SOURCE FILE LIST
# ============================================================

list_frame = tk.Frame(
    source_frame
)

list_frame.pack(
    fill="both",
    expand=True
)


source_listbox = tk.Listbox(
    list_frame,
    width=100,
    height=8
)

source_listbox.pack(
    side="left",
    fill="both",
    expand=True
)


scrollbar = tk.Scrollbar(
    list_frame,
    orient="vertical",
    command=source_listbox.yview
)

scrollbar.pack(
    side="right",
    fill="y"
)


source_listbox.config(
    yscrollcommand=scrollbar.set
)


# ============================================================
# CONSOLIDATED FILE FRAME
# ============================================================

consolidated_frame = tk.LabelFrame(
    root,
    text="Consolidated Excel File",
    padx=10,
    pady=10
)

consolidated_frame.pack(
    fill="x",
    padx=30,
    pady=10
)


tk.Entry(
    consolidated_frame,
    textvariable=consolidated_var,
    width=75
).pack(
    side="left",
    padx=5
)


tk.Button(
    consolidated_frame,
    text="Open Existing",
    command=choose_consolidated,
    width=15
).pack(
    side="left",
    padx=5
)


tk.Button(
    consolidated_frame,
    text="Create New",
    command=create_new_consolidated,
    width=15
).pack(
    side="left",
    padx=5
)


# ============================================================
# MAIN ACTION BUTTON
# ============================================================

tk.Button(
    root,
    text="EXTRACT & CONSOLIDATE ALL FILES",
    command=run_consolidation,
    font=(
        "Arial",
        12,
        "bold"
    ),
    width=35,
    height=2
).pack(
    pady=15
)


# ============================================================
# FOOTER
# ============================================================

tk.Label(
    root,
    text=(
        "Supported: Excel (.xlsx, .xls) and PDF (.pdf) | "
        "Duplicate Invoice Numbers are automatically skipped."
    ),
    fg="gray"
).pack(
    pady=5
)


# ============================================================
# START APPLICATION
# ============================================================

root.mainloop()
```

