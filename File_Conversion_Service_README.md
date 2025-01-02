
# File Conversion Service

This is a Flask-based web application that provides file conversion services between Microsoft Word (.docx) and PDF file formats. The application includes:

- A user interface for uploading files.
- Backend functionality for converting Word documents to PDFs and vice versa.

## Features

1. **Convert Word to PDF:**
   - Upload a Word document (.docx), and the application will convert it into a PDF.

2. **Convert PDF to Word:**
   - Upload a PDF, and the application will convert it into a Word document (.docx).

## Prerequisites

To run this application, ensure you have the following installed:

- Python 3.8 or later
- Flask
- comtypes (for interacting with Microsoft Office)
- Microsoft Word (required for Word-to-PDF conversion)

## Installation

1. Clone the repository or download the project files.

   ```bash
   git clone https://github.com/your-repo/file-conversion-service.git
   cd file-conversion-service
   ```

2. Create a virtual environment to isolate dependencies:

   ```bash
   python -m venv env
   source env/bin/activate  # On Windows, use `env\Scripts\activate`
   ```

3. Install the required Python packages:

   ```bash
   pip install -r requirements.txt
   ```

4. Ensure that the necessary directories exist for uploads and downloads:

   - `uploads/`: Stores user-uploaded files temporarily.
   - `downloads/`: Stores converted files for download.

   These directories will be automatically created if they do not exist.

## Usage

1. Start the Flask application:

   ```bash
   python app.py
   ```

2. Open your web browser and navigate to:

   ```
   http://127.0.0.1:5000/
   ```

3. Use the interface to upload and convert files:
   - Choose a Word document to convert to PDF.
   - Choose a PDF file to convert to a Word document.

4. Download the converted files directly from the interface.

## Conversion Details

- **Word to PDF Conversion:**
  The `comtypes` library is used to interact with Microsoft Word's COM interface for converting documents.

- **File Processing:**
  - Files are sanitized to prevent potential security risks.
  - Temporary files are stored in the `uploads/` directory.
  - Converted files are saved in the `downloads/` directory.

## Error Handling

- If a conversion fails, the application returns a 500 error with an appropriate message.
- Ensure Microsoft Word is properly installed and configured to avoid runtime errors.

## Known Limitations

- Requires Microsoft Word for Word-to-PDF conversions.
- The application is designed for local or limited network use and may not be production-ready without additional security and scalability features.

## License

This project is licensed under the MIT License. See the `LICENSE` file for details.

## Contributing

Contributions are welcome! Please submit a pull request or open an issue for bug reports and feature requests.

## Contact

For questions or support, please contact:

- **Author:** Your Name
- **Email:** your.email@example.com
