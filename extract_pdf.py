import pdfplumber
import pypdf
import os

pdf_path = "EXSIM Case.pdf"
output_path = "extracted_case_data.txt"

def extract_pdf_data(pdf_path, output_path):
    print(f"Starting extraction from {pdf_path}...")
    
    extracted_text = []
    
    # 1. Extract Text using pypdf for raw text flow
    print("Extracting raw text with pypdf...")
    try:
        reader = pypdf.PdfReader(pdf_path)
        for i, page in enumerate(reader.pages):
            text = page.extract_text()
            extracted_text.append(f"--- Page {i+1} (pypdf) ---\n{text}\n")
    except Exception as e:
        extracted_text.append(f"Error extracting text with pypdf: {e}\n")

    # 2. Extract Tables and Text using pdfplumber for layout preservation
    print("Extracting tables/layout with pdfplumber...")
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for i, page in enumerate(pdf.pages):
                extracted_text.append(f"--- Page {i+1} (pdfplumber) ---\n")
                
                # Extract text
                text = page.extract_text()
                extracted_text.append(f"*** Text ***\n{text}\n")
                
                # Extract tables
                tables = page.extract_tables()
                if tables:
                    extracted_text.append("*** Tables ***\n")
                    for table in tables:
                        for row in table:
                            # Clean None values and join
                            row_str = " | ".join([str(cell) if cell is not None else "" for cell in row])
                            extracted_text.append(f"{row_str}\n")
                        extracted_text.append("\n[End of Table]\n")
                extracted_text.append("\n")
                
    except Exception as e:
        extracted_text.append(f"Error extracting with pdfplumber: {e}\n")

    # Save to file
    print(f"Saving extracted data to {output_path}...")
    with open(output_path, "w", encoding="utf-8") as f:
        f.writelines(extracted_text)
    
    print("Extraction complete.")

if __name__ == "__main__":
    if os.path.exists(pdf_path):
        extract_pdf_data(pdf_path, output_path)
    else:
        print(f"Error: File {pdf_path} not found.")
