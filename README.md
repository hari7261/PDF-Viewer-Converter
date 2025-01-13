# Enhanced PDF Viewer & Converter

## Overview


The **Enhanced PDF Viewer & Converter** is a desktop application designed to provide a user-friendly interface for viewing, navigating, and converting PDF and Word documents. Built using Python and PyQt5, this application combines the power of document manipulation with an intuitive graphical interface, making it easy to manage PDFs and Word files.

---

## Features

### 1. **PDF Viewer**
   - **Page Navigation:**  
     - Easily navigate through pages using **Previous Page** and **Next Page** buttons.
     - Directly jump to a specific page by entering the page number in the **Page Spin Box**.
   - **Zoom Controls:**  
     - Use the **Zoom In** and **Zoom Out** buttons or the **Zoom Slider** to adjust the zoom level from 50% to 300%. 
     - The current zoom percentage is displayed in real-time.
   - **Smooth Scrolling:**  
     - A scrollable area lets you easily view large pages with smooth horizontal and vertical scrolling.

### 2. **PDF to Word Conversion**
   - Convert a PDF document to a Word file (DOCX format) using the `pdf2docx` library.
   - Steps:
     1. Select a PDF file to convert.
     2. Specify the output Word file name and location.
     3. The application processes the file and saves the Word document.
   - Includes error handling for unsupported or corrupted PDF files.

### 3. **Word to PDF Conversion**
   - Convert Word documents (DOCX/DOC) into PDF format using the `docx2pdf` library.
   - Steps:
     1. Select a Word file to convert.
     2. Specify the output PDF file name and location.
     3. The application processes the file and saves the PDF.
   - Supports a variety of Word document formats.

### 4. **Smooth and High-Quality Rendering**
   - Uses the `PyMuPDF` (fitz) library to render PDF pages with high clarity.
   - Pages are rendered at a high zoom factor (2.5x) for crisp and readable text.
   - Automatically scales rendered pages according to the zoom level for consistent quality.

### 5. **Responsive and Modern User Interface**
   - Built using PyQt5 with a focus on ease of use and aesthetic design.
   - **Buttons and Controls Styling:**  
     - Modern button styles with hover effects for better user experience.
   - **Adaptive Layouts:**  
     - The main window automatically adjusts to the screen size and maximizes for optimal viewing.

---

## How to Use the Application

### 1. **Open a PDF File**
   - Click the **Open PDF** button.
   - Select a PDF file from your system.
   - The application will load the file and display the first page.
   - Use navigation and zoom controls to view the document.

### 2. **Navigate Through Pages**
   - Use the **Previous Page** and **Next Page** buttons to move one page at a time.
   - Enter the desired page number in the spin box to jump to a specific page.

### 3. **Zoom In and Out**
   - Use the **Zoom In** or **Zoom Out** buttons to adjust the zoom level.
   - Alternatively, use the **Zoom Slider** to set the zoom percentage directly.

### 4. **Convert PDF to Word**
   - Click the **Convert PDF to Word** button.
   - Select a PDF file and specify the output location for the Word document.
   - Wait for the conversion to complete, and a success message will appear.

### 5. **Convert Word to PDF**
   - Click the **Convert Word to PDF** button.
   - Select a Word file and specify the output location for the PDF document.
   - Wait for the conversion to complete, and a success message will appear.

---

## Technical Details

- **Core Libraries Used:**
  - `PyQt5`: For the graphical user interface.
  - `fitz` (PyMuPDF): For rendering PDF pages and high-quality image generation.
  - `pdf2docx`: For converting PDFs to Word documents.
  - `docx2pdf`: For converting Word documents to PDF files.
  - `numpy`: For image data manipulation.

- **High-Performance Rendering:**  
  - Renders pages using a zoom matrix (`fitz.Matrix`) to ensure sharpness and clarity.
  - Converts RGBA images to RGB for optimal display.

- **User-Friendly Controls:**
  - The interface is intuitive and minimizes user effort.
  - Error messages are displayed for invalid inputs or unsupported file formats.

---

## Installation and Requirements

### Prerequisites:
- **Python 3.8 or later**
- Required Python Libraries:
  - `PyQt5`
  - `fitz` (PyMuPDF)
  - `pdf2docx`
  - `docx2pdf`
  - `numpy`

### Installation:
1. Install the required libraries using pip:
   ```bash
   pip install PyQt5 fitz pdf2docx docx2pdf numpy
   ```
2. Run the application:
   ```bash
   python your_script_name.py
   ```

---
