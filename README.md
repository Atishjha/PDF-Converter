# PDF-Converter
This code creates a graphical user interface (GUI) application that allows users to:

Select an input file (Word document, Excel spreadsheet, text file, or image)
Specify an output PDF file location
Convert the input file to PDF format

Main Components
Libraries Used

Apache PDFBox: For PDF creation and manipulation
Apache POI: For working with Microsoft Office documents (Word and Excel)
Swing/AWT: For the graphical user interface

Class Structure
The application is built as a single class called PDFConverter with the following key components:

User Interface Elements:

Text fields for input and output file paths
Browse buttons for file selection
File type dropdown for selecting conversion type
Convert button to start the process
Status label and progress bar


Conversion Methods:

One method for each supported file type (Word, Excel, Text, Image)
Helper methods for specific file formats (DOCX vs DOC)



Detailed Breakdown
Main Method
The entry point creates an instance of the PDFConverter and displays the interface:
javapublic static void main(String[] args) {
    EventQueue.invokeLater(() -> {
        try {
            PDFConverter window = new PDFConverter();
            window.frame.setVisible(true);
        } catch (Exception e) {
            JOptionPane.showMessageDialog(null, "Failed to initialize application: " + e.getMessage(),
                    "Error", JOptionPane.ERROR_MESSAGE);
        }
    });
}
UI Initialization
The initialize() method sets up the user interface with a grid layout that includes:

Labels and text fields for input/output files
Browse buttons for file selection
File type selector dropdown
Convert button and status indicators

Action Listeners
The setupActionListeners() method connects user actions to behaviors:

Input Browse Button:

Opens file chooser dialog
Filters files based on selected type
Sets the input file path in the text field
Auto-generates an output file name


Output Browse Button:

Opens file save dialog
Proposes output filename based on input
Confirms before overwriting existing files


Convert Button:

Validates input/output files
Disables UI during conversion
Runs conversion in background thread
Updates UI with success/failure status



Conversion Methods
Word to PDF (convertWordToPdf)
Handles both DOC and DOCX formats:

For DOCX: Uses XWPFDocument to extract paragraphs
For DOC: Uses HWPFDocument with WordExtractor
Creates PDF pages and adds text content
Handles pagination when text exceeds page limits

Excel to PDF (convertExcelToPdf)

Extracts data from the first sheet
Creates a table-like layout in the PDF
Handles different cell types (string, numeric, boolean)
Supports pagination for large spreadsheets

Text to PDF (convertTextToPdf)

Reads all lines from the text file
Creates PDF with appropriate font (Courier)
Handles line wrapping for long text
Creates new pages when content exceeds page limits

Image to PDF (convertImageToPdf)

Loads image using PDImageXObject
Scales image to fit on PDF page while maintaining aspect ratio
Centers the image on the page

Helper Methods

getRootCauseMessage: Extracts meaningful error messages
setUIEnabled: Enables/disables UI components during processing
getMaxColumns: Determines the number of columns to display from Excel

Key Implementation Details
Error Handling

Comprehensive try-catch blocks
User-friendly error messages
Error message extraction from root causes

Multi-threading

Conversion runs in a background thread via SwingWorker
UI remains responsive during conversion
Progress updates are safely dispatched to the UI thread

Resource Management

All streams and documents are properly closed using try-with-resources
PDFs are properly saved and closed

User Experience

Auto-populated output filenames
Overwrite confirmations
File filters for appropriate selections
Status updates and progress indication

Limitations

Basic text formatting - it doesn't preserve complex formatting from Word documents
Only processes the first sheet of Excel files
Limited table formatting in Excel conversions
No support for Word tables, images, or complex elements
Image conversion is limited to basic scaling

This application provides a functional desktop tool for basic file conversions to PDF format with a user-friendly interface and robust error handling.RetryClaude does not have the ability to run the code it generates yet.Claude can make mistakes. Please double-check responses.
