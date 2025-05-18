import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.List;

public class PDFConverter {
    private JFrame frame;
    private JTextField inputFileField;
    private JTextField outputFileField;
    private JButton browseInputButton;
    private JButton browseOutputButton;
    private JButton convertButton;
    private JLabel statusLabel;
    private JProgressBar progressBar;
    private JComboBox<String> fileTypeComboBox;

    public static void main(String[] args) {
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

    public PDFConverter() {
        initialize();
    }

    private void initialize() {
        frame = new JFrame("PDF Converter");
        frame.setBounds(100, 100, 550, 300);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.getContentPane().setLayout(new GridBagLayout());
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(5, 5, 5, 5);
        gbc.fill = GridBagConstraints.HORIZONTAL;

        // Input File Components
        gbc.gridx = 0;
        gbc.gridy = 0;
        frame.getContentPane().add(new JLabel("Input File:"), gbc);

        gbc.gridx = 1;
        gbc.weightx = 1;
        inputFileField = new JTextField();
        frame.getContentPane().add(inputFileField, gbc);

        gbc.gridx = 2;
        gbc.weightx = 0;
        browseInputButton = new JButton("Browse");
        frame.getContentPane().add(browseInputButton, gbc);

        // Output File Components
        gbc.gridx = 0;
        gbc.gridy = 1;
        frame.getContentPane().add(new JLabel("Output File:"), gbc);

        gbc.gridx = 1;
        gbc.weightx = 1;
        outputFileField = new JTextField();
        frame.getContentPane().add(outputFileField, gbc);

        gbc.gridx = 2;
        gbc.weightx = 0;
        browseOutputButton = new JButton("Browse");
        frame.getContentPane().add(browseOutputButton, gbc);

        // File Type Components
        gbc.gridx = 0;
        gbc.gridy = 2;
        frame.getContentPane().add(new JLabel("File Type:"), gbc);

        gbc.gridx = 1;
        gbc.gridwidth = 2;
        fileTypeComboBox = new JComboBox<>(new String[]{"Word (.doc, .docx)", "Excel (.xlsx)", "Text (.txt)", "Image (.jpg, .png)"});
        frame.getContentPane().add(fileTypeComboBox, gbc);

        // Convert Button
        gbc.gridx = 1;
        gbc.gridy = 3;
        gbc.gridwidth = 1;
        convertButton = new JButton("Convert");
        frame.getContentPane().add(convertButton, gbc);

        // Status Label
        gbc.gridx = 0;
        gbc.gridy = 4;
        gbc.gridwidth = 3;
        statusLabel = new JLabel("Status: Ready");
        frame.getContentPane().add(statusLabel, gbc);

        // Progress Bar
        gbc.gridy = 5;
        progressBar = new JProgressBar();
        frame.getContentPane().add(progressBar, gbc);

        setupActionListeners();
    }

    private void setupActionListeners() {
        browseInputButton.addActionListener(e -> {
            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setDialogTitle("Select Input File");

            int selection = fileTypeComboBox.getSelectedIndex();
            switch (selection) {
                case 0: // Word
                    fileChooser.setFileFilter(new FileNameExtensionFilter("Word Documents", "doc", "docx"));
                    break;
                case 1: // Excel
                    fileChooser.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx"));
                    break;
                case 2: // Text
                    fileChooser.setFileFilter(new FileNameExtensionFilter("Text Files", "txt"));
                    break;
                case 3: // Image
                    fileChooser.setFileFilter(new FileNameExtensionFilter("Image Files", "jpg", "jpeg", "png"));
                    break;
            }

            int result = fileChooser.showOpenDialog(frame);
            if (result == JFileChooser.APPROVE_OPTION) {
                File selectedFile = fileChooser.getSelectedFile();
                inputFileField.setText(selectedFile.getAbsolutePath());

                // Auto-generate output file name
                String inputPath = selectedFile.getAbsolutePath();
                String outputPath = inputPath.substring(0, inputPath.lastIndexOf('.')) + ".pdf";
                outputFileField.setText(outputPath);
            }
        });

        browseOutputButton.addActionListener(e -> {
            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setDialogTitle("Save PDF File");
            fileChooser.setFileFilter(new FileNameExtensionFilter("PDF Files", "pdf"));

            // Set suggested file name if input file exists
            if (!inputFileField.getText().isEmpty()) {
                File inputFile = new File(inputFileField.getText());
                if (inputFile.exists()) {
                    String name = inputFile.getName();
                    name = name.substring(0, name.lastIndexOf('.')) + ".pdf";
                    fileChooser.setSelectedFile(new File(inputFile.getParent(), name));
                }
            }

            int result = fileChooser.showSaveDialog(frame);
            if (result == JFileChooser.APPROVE_OPTION) {
                File selectedFile = fileChooser.getSelectedFile();
                String path = selectedFile.getAbsolutePath();
                if (!path.toLowerCase().endsWith(".pdf")) {
                    path += ".pdf";
                }

                // Check if file exists and confirm overwrite
                if (new File(path).exists()) {
                    int confirm = JOptionPane.showConfirmDialog(frame,
                            "The file already exists. Do you want to overwrite it?",
                            "Confirm Overwrite", JOptionPane.YES_NO_OPTION);
                    if (confirm != JOptionPane.YES_OPTION) {
                        return;
                    }
                }

                outputFileField.setText(path);
            }
        });

        convertButton.addActionListener(e -> {
            String inputFile = inputFileField.getText();
            String outputFile = outputFileField.getText();

            // Validate inputs
            if (inputFile.isEmpty() || outputFile.isEmpty()) {
                statusLabel.setText("Status: Please select input and output files.");
                JOptionPane.showMessageDialog(frame, "Please select both input and output files",
                        "Error", JOptionPane.ERROR_MESSAGE);
                return;
            }

            File input = new File(inputFile);
            if (!input.exists()) {
                statusLabel.setText("Status: Input file does not exist.");
                JOptionPane.showMessageDialog(frame, "Input file does not exist",
                        "Error", JOptionPane.ERROR_MESSAGE);
                return;
            }

            // Disable UI during conversion
            setUIEnabled(false);
            statusLabel.setText("Status: Converting...");
            progressBar.setIndeterminate(true);

            // Run conversion in background thread
            new SwingWorker<Void, Void>() {
                @Override
                protected Void doInBackground() throws Exception {
                    try {
                        int fileType = fileTypeComboBox.getSelectedIndex();
                        convertToPdf(inputFile, outputFile, fileType);
                        SwingUtilities.invokeLater(() -> {
                            statusLabel.setText("Status: Conversion completed successfully!");
                            progressBar.setIndeterminate(false);
                            progressBar.setValue(100);
                            JOptionPane.showMessageDialog(frame, "File converted successfully!",
                                    "Success", JOptionPane.INFORMATION_MESSAGE);
                        });
                    } catch (Exception ex) {
                        SwingUtilities.invokeLater(() -> {
                            statusLabel.setText("Status: Error - " + getRootCauseMessage(ex));
                            progressBar.setIndeterminate(false);
                            progressBar.setValue(0);
                            JOptionPane.showMessageDialog(frame,
                                    "Error during conversion: " + getRootCauseMessage(ex),
                                    "Error", JOptionPane.ERROR_MESSAGE);
                        });
                    } finally {
                        SwingUtilities.invokeLater(() -> setUIEnabled(true));
                    }
                    return null;
                }
            }.execute();
        });
    }

    private String getRootCauseMessage(Throwable t) {
        Throwable rootCause = t;
        while (rootCause.getCause() != null && rootCause.getCause() != rootCause) {
            rootCause = rootCause.getCause();
        }
        return rootCause.getMessage() != null ? rootCause.getMessage() : "Unknown error";
    }

    private void setUIEnabled(boolean enabled) {
        convertButton.setEnabled(enabled);
        browseInputButton.setEnabled(enabled);
        browseOutputButton.setEnabled(enabled);
        fileTypeComboBox.setEnabled(enabled);
    }

    private void convertToPdf(String inputFile, String outputFile, int fileType) throws IOException {
        switch (fileType) {
            case 0: // Word
                convertWordToPdf(inputFile, outputFile);
                break;
            case 1: // Excel
                convertExcelToPdf(inputFile, outputFile);
                break;
            case 2: // Text
                convertTextToPdf(inputFile, outputFile);
                break;
            case 3: // Image
                convertImageToPdf(inputFile, outputFile);
                break;
            default:
                throw new IllegalArgumentException("Unsupported file type");
        }
    }

    private void convertWordToPdf(String inputFile, String outputFile) throws IOException {
        try (PDDocument document = new PDDocument()) {
            if (inputFile.toLowerCase().endsWith(".docx")) {
                processDocxDocument(document, inputFile);
            } else if (inputFile.toLowerCase().endsWith(".doc")) {
                processDocDocument(document, inputFile);
            } else {
                throw new IOException("Unsupported Word file format");
            }
            document.save(outputFile);
        }
    }

    private void processDocxDocument(PDDocument document, String inputFile) throws IOException {
        try (FileInputStream fis = new FileInputStream(inputFile);
             XWPFDocument docx = new XWPFDocument(fis)) {

            List<XWPFParagraph> paragraphs = docx.getParagraphs();
            if (paragraphs.isEmpty()) {
                throw new IOException("The document contains no text");
            }

            PDPage currentPage = new PDPage(PDRectangle.A4);
            document.addPage(currentPage);

            PDPageContentStream currentStream = null;
            try {
                currentStream = new PDPageContentStream(document, currentPage);
                currentStream.beginText();
                currentStream.setFont(PDType1Font.HELVETICA, 12);
                currentStream.newLineAtOffset(50, 750);

                float yPosition = 750;
                for (XWPFParagraph para : paragraphs) {
                    String text = para.getText();
                    if (text == null || text.trim().isEmpty()) {
                        continue;
                    }

                    if (yPosition < 50) {
                        // End current stream and page
                        currentStream.endText();
                        currentStream.close();
                        currentStream = null;

                        // Create new page and stream
                        currentPage = new PDPage(PDRectangle.A4);
                        document.addPage(currentPage);
                        currentStream = new PDPageContentStream(document, currentPage);
                        currentStream.beginText();
                        currentStream.setFont(PDType1Font.HELVETICA, 12);
                        yPosition = 750;
                        currentStream.newLineAtOffset(50, yPosition);
                    }

                    currentStream.showText(text);
                    yPosition -= 15;
                    currentStream.newLineAtOffset(0, -15);
                }

                // Make sure to end the text operation
                if (currentStream != null) {
                    currentStream.endText();
                }
            } finally {
                // Make sure stream is always closed
                if (currentStream != null) {
                    currentStream.close();
                }
            }
        }
    }

    private void processDocDocument(PDDocument document, String inputFile) throws IOException {
        try (FileInputStream fis = new FileInputStream(inputFile);
             HWPFDocument doc = new HWPFDocument(fis);
             WordExtractor extractor = new WordExtractor(doc)) {

            String[] paragraphs = extractor.getParagraphText();
            if (paragraphs == null || paragraphs.length == 0) {
                throw new IOException("The document contains no text");
            }

            PDPage currentPage = new PDPage(PDRectangle.A4);
            document.addPage(currentPage);

            PDPageContentStream currentStream = null;
            try {
                currentStream = new PDPageContentStream(document, currentPage);
                currentStream.beginText();
                currentStream.setFont(PDType1Font.HELVETICA, 12);
                currentStream.newLineAtOffset(50, 750);

                float yPosition = 750;
                for (String paragraph : paragraphs) {
                    if (paragraph == null || paragraph.trim().isEmpty()) {
                        continue;
                    }

                    paragraph = paragraph.trim().replace('\r', ' ').replace('\n', ' ');

                    if (yPosition < 50) {
                        currentStream.endText();
                        currentStream.close();
                        currentStream = null;

                        currentPage = new PDPage(PDRectangle.A4);
                        document.addPage(currentPage);
                        currentStream = new PDPageContentStream(document, currentPage);
                        currentStream.beginText();
                        currentStream.setFont(PDType1Font.HELVETICA, 12);
                        yPosition = 750;
                        currentStream.newLineAtOffset(50, yPosition);
                    }

                    currentStream.showText(paragraph);
                    yPosition -= 15;
                    currentStream.newLineAtOffset(0, -15);
                }

                // Make sure to end the text operation
                if (currentStream != null) {
                    currentStream.endText();
                }
            } finally {
                // Make sure stream is always closed
                if (currentStream != null) {
                    currentStream.close();
                }
            }
        }
    }

    private void convertExcelToPdf(String inputFile, String outputFile) throws IOException {
        try (FileInputStream fis = new FileInputStream(inputFile);
             XSSFWorkbook workbook = new XSSFWorkbook(fis);
             PDDocument document = new PDDocument()) {

            if (workbook.getNumberOfSheets() == 0) {
                throw new IOException("The workbook contains no sheets");
            }

            XSSFSheet sheet = workbook.getSheetAt(0);
            if (sheet.getPhysicalNumberOfRows() == 0) {
                throw new IOException("The sheet contains no data");
            }

            PDPage page = new PDPage(PDRectangle.A4);
            document.addPage(page);
            PDPageContentStream contentStream = null;

            try {
                contentStream = new PDPageContentStream(document, page);
                contentStream.setFont(PDType1Font.HELVETICA, 10);

                float margin = 50;
                float yStart = page.getMediaBox().getHeight() - margin;
                float tableWidth = page.getMediaBox().getWidth() - 2 * margin;
                float yPosition = yStart;
                float rowHeight = 20f;
                int maxRowsPerPage = (int) ((yStart - margin) / rowHeight);
                int rowCounter = 0;

                int maxCols = getMaxColumns(sheet);
                float colWidth = tableWidth / maxCols;

                for (Row row : sheet) {
                    if (row == null) continue;

                    if (rowCounter >= maxRowsPerPage) {
                        contentStream.close();
                        contentStream = null;

                        page = new PDPage(PDRectangle.A4);
                        document.addPage(page);
                        contentStream = new PDPageContentStream(document, page);
                        contentStream.setFont(PDType1Font.HELVETICA, 10);
                        yPosition = yStart;
                        rowCounter = 0;
                    }

                    float xPosition = margin;

                    for (int c = 0; c < maxCols; c++) {
                        contentStream.beginText();
                        contentStream.newLineAtOffset(xPosition, yPosition);

                        Cell cell = row.getCell(c, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                        String text = "";

                        if (cell.getCellType() == CellType.STRING) {
                            text = cell.getStringCellValue();
                        } else if (cell.getCellType() == CellType.NUMERIC) {
                            text = String.format("%.2f", cell.getNumericCellValue());
                        } else if (cell.getCellType() == CellType.BOOLEAN) {
                            text = String.valueOf(cell.getBooleanCellValue());
                        }

                        if (text.length() > 15) {
                            text = text.substring(0, 12) + "...";
                        }

                        contentStream.showText(text);
                        contentStream.endText(); // Properly end text for each cell
                        xPosition += colWidth;
                    }

                    rowCounter++;
                    yPosition -= rowHeight;
                }
            } finally {
                if (contentStream != null) {
                    contentStream.close();
                }
            }

            document.save(outputFile);
        }
    }

    private int getMaxColumns(XSSFSheet sheet) {
        int maxColumns = 0;
        for (Row row : sheet) {
            if (row.getLastCellNum() > maxColumns) {
                maxColumns = row.getLastCellNum();
            }
        }
        return Math.min(maxColumns, 15);
    }

    private void convertTextToPdf(String inputFile, String outputFile) throws IOException {
        try (PDDocument document = new PDDocument()) {
            PDPage currentPage = new PDPage(PDRectangle.A4);
            document.addPage(currentPage);

            List<String> lines = Files.readAllLines(Paths.get(inputFile), StandardCharsets.UTF_8);
            if (lines.isEmpty()) {
                throw new IOException("The text file is empty");
            }

            // Initialize first content stream
            PDPageContentStream currentStream = null;
            try {
                currentStream = new PDPageContentStream(document, currentPage);
                currentStream.beginText();
                currentStream.setFont(PDType1Font.COURIER, 12);
                currentStream.newLineAtOffset(50, 750);

                float leading = 14;
                float y = 750;

                for (String line : lines) {
                    if (y < 50) {
                        // End current stream and page
                        currentStream.endText();
                        currentStream.close();
                        currentStream = null;

                        // Create new page and stream
                        currentPage = new PDPage(PDRectangle.A4);
                        document.addPage(currentPage);
                        currentStream = new PDPageContentStream(document, currentPage);
                        currentStream.beginText();
                        currentStream.setFont(PDType1Font.COURIER, 12);
                        y = 750;
                        currentStream.newLineAtOffset(50, y);
                    }

                    // Handle long lines by splitting them
                    if (line.length() > 100) {
                        String[] parts = line.split("(?<=\\G.{100})");
                        for (String part : parts) {
                            if (y < 50) {
                                currentStream.endText();
                                currentStream.close();
                                currentStream = null;

                                currentPage = new PDPage(PDRectangle.A4);
                                document.addPage(currentPage);
                                currentStream = new PDPageContentStream(document, currentPage);
                                currentStream.beginText();
                                currentStream.setFont(PDType1Font.COURIER, 12);
                                y = 750;
                                currentStream.newLineAtOffset(50, y);
                            }
                            currentStream.showText(part);
                            y -= leading;
                            currentStream.newLineAtOffset(0, -leading);
                        }
                    } else {
                        currentStream.showText(line);
                        y -= leading;
                        currentStream.newLineAtOffset(0, -leading);
                    }
                }

                // Make sure to end the text operation
                if (currentStream != null) {
                    currentStream.endText();
                }
            } finally {
                // Make sure stream is always closed
                if (currentStream != null) {
                    currentStream.close();
                }
            }

            document.save(outputFile);
        }
    }

    private void convertImageToPdf(String inputFile, String outputFile) throws IOException {
        try (PDDocument document = new PDDocument()) {
            PDPage page = new PDPage(PDRectangle.A4);
            document.addPage(page);

            PDImageXObject image;
            try {
                image = PDImageXObject.createFromFile(inputFile, document);
            } catch (IOException e) {
                throw new IOException("Unsupported image format or corrupted image file", e);
            }

            try (PDPageContentStream contentStream = new PDPageContentStream(document, page)) {
                float pageWidth = page.getMediaBox().getWidth();
                float pageHeight = page.getMediaBox().getHeight();
                float margin = 50;

                float availableWidth = pageWidth - 2 * margin;
                float availableHeight = pageHeight - 2 * margin;

                float imgWidth = image.getWidth();
                float imgHeight = image.getHeight();

                float scaleW = availableWidth / imgWidth;
                float scaleH = availableHeight / imgHeight;
                float scale = Math.min(scaleW, scaleH);

                float scaledWidth = imgWidth * scale;
                float scaledHeight = imgHeight * scale;

                float x = margin + (availableWidth - scaledWidth) / 2;
                float y = margin + (availableHeight - scaledHeight) / 2;

                contentStream.drawImage(image, x, y, scaledWidth, scaledHeight);
            }

            document.save(outputFile);
        }
    }
}