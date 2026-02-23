package com.arisglobal.reads3files;

import com.amazonaws.auth.DefaultAWSCredentialsProviderChain;
import com.amazonaws.services.s3.AmazonS3;
import com.amazonaws.services.s3.AmazonS3ClientBuilder;
import com.amazonaws.services.s3.model.S3Object;
import com.amazonaws.services.s3.model.S3ObjectInputStream;
import com.arisglobal.reads3files.components.AppConfigs;
import com.arisglobal.reads3files.service.AsposeLicenseHandler;
import com.arisglobal.reads3files.service.S3Configurations;
import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.Row;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.pdf.Document;
import com.aspose.pdf.Field;
import com.aspose.pdf.FileSpecification;
import com.aspose.pdf.TextBoxField;
import com.aspose.words.FormField;
import com.aspose.words.HeightRule;
import com.aspose.words.HorizontalAlignment;
import com.aspose.words.LayoutCollector;
import com.aspose.words.LayoutEntityType;
import com.aspose.words.LayoutEnumerator;
import com.aspose.words.Node;
import com.aspose.words.NodeCollection;
import com.aspose.words.NodeType;
import com.aspose.words.Paragraph;
import com.aspose.words.Run;
import com.aspose.words.SaveFormat;
import com.aspose.words.Section;
import com.aspose.words.StructuredDocumentTag;
import com.aspose.words.Table;
import com.aspose.words.TextWrapping;
import jakarta.annotation.PostConstruct;
import jakarta.mail.BodyPart;
import jakarta.mail.Message;
import jakarta.mail.Multipart;
import jakarta.mail.Part;
import jakarta.mail.internet.MimeMessage;
import jakarta.mail.internet.MimeUtility;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.Setter;
import lombok.extern.slf4j.Slf4j;
import net.lingala.zip4j.ZipFile;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.io.IOUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.tika.Tika;
import org.apache.tika.metadata.Metadata;
import org.springframework.context.ApplicationContext;
import org.springframework.core.io.ClassPathResource;
import org.springframework.stereotype.Component;
import org.springframework.stereotype.Service;

import java.awt.*;
import java.awt.font.FontRenderContext;
import java.awt.geom.AffineTransform;
import java.awt.geom.Rectangle2D;
import java.io.BufferedInputStream;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

@AllArgsConstructor
@Getter
@Setter
@Component
@Slf4j
@Service
public class WritePageNumbers {

    private final ApplicationContext applicationContext;
    private final AppConfigs appConfigs;
    private final boolean isLocal = false;
    private final S3Configurations s3Configurations;
    public final String ADOBE_LIVECYCLE_DESIGNER = "adobe livecycle designer";


    @PostConstruct
    public void loadAsposeLicense() {
        AsposeLicenseHandler licenseHandler = new AsposeLicenseHandler();
        /**
         * This code is to Enable the license for Aspose.
         */
        try {
            licenseHandler.initialiseAsposeLicense(new ClassPathResource("Aspose.Total.Java.lic").getInputStream());
        } catch (Exception e) {
            log.error("caught Exception initialiseAsposeLicense", e);
        }

        try {
            licenseHandler.initialiseAsposeLicenseForExcel(new ClassPathResource("Aspose.Total.Java.lic").getInputStream());
        } catch (Exception e) {
            log.error("caught Exception initialiseAsposeLicenseForExcel", e);
        }

        try {
            licenseHandler.initialiseAsposeLicenseForPDF(new ClassPathResource("Aspose.Total.Java.lic").getInputStream());
        } catch (Exception e) {
            log.error("caught Exception initialiseAsposeLicenseForPDF", e);
        }


        try {
            licenseHandler.initialiseAsposeLicenseForPDFKit(new ClassPathResource("Aspose.Total.Java.lic").getInputStream());
        } catch (Exception e) {
            log.error("caught Exception initialiseAsposeLicenseForPDF", e);
        }
    }

    private boolean isMultilineWordDoc(com.aspose.words.Document document, Detail detail) throws Exception {
        for (FormField field : (Iterable<FormField>)
                document.getChildNodes(NodeType.FORM_FIELD, true)) {
            if (StringUtils.isNotBlank(field.getResult()) && isFormFieldOverflowDetected(document, field, detail)) {
                return true;
            }
        }
        return false;
    }

    public boolean isFormFieldOverflowDetected(com.aspose.words.Document doc, StructuredDocumentTag field) throws Exception {
        doc.updatePageLayout();
        com.aspose.words.Row row = (com.aspose.words.Row) field.getAncestor(NodeType.ROW);
        boolean isPara = false;
        boolean isOverflow = false;
        if (row != null) {
            if (row.getRowFormat().getHeightRule() == HeightRule.EXACTLY) {
                // Approach 1: Check if the form field is inside a paragraph and has an associated layout entity
                Paragraph para = (Paragraph) field.getAncestor(NodeType.PARAGRAPH);
                doc.updatePageLayout();
                LayoutCollector collector = new LayoutCollector(doc);
                LayoutEnumerator enumerator = new LayoutEnumerator(doc);
                Object entity = collector.getEntity(para);
                isPara = entity != null;

                if (isPara) {
                    enumerator.setCurrent(collector.getEntity(para));
                    double totalHeight = 0;
                    if (enumerator.moveFirstChild()) {
                        do {
                            if (enumerator.getType() == LayoutEntityType.LINE) {
                                Rectangle2D rect = enumerator.getRectangle();
                                totalHeight += rect.getHeight();
                            }
                        } while (enumerator.moveNext());
                    }
                }

                // Approach 2: Check if the form field's layout rectangle exceeds the row height
                enumerator.setCurrent(collector.getEntity(field));
                Rectangle2D rect = enumerator.getRectangle();
                double renderedHeight = rect.getHeight();
                double rowHeight = row.getRowFormat().getHeight();
                if (renderedHeight > rowHeight)
                    isOverflow = true;
            }
        }
        return isPara && isOverflow;
    }

    public boolean isFormFieldOverflowDetected(com.aspose.words.Document doc, FormField field, Detail detail) throws Exception {
        boolean isOverflow = false;
        try {
            doc.updatePageLayout();
            com.aspose.words.Row row = (com.aspose.words.Row) field.getAncestor(NodeType.ROW);
            if (row != null) {
                if (row.getRowFormat().getHeightRule() == HeightRule.EXACTLY) {
                    // Approach 3: Check if text height is more than rectangle height
                    String text = field.getResult();
                    Run run = getFirstResultRun(field);
                    if (null != run) {
                        String fontName = run.getFont().getName();
                        int fontSize = (int) run.getFont().getSize();
                        com.aspose.words.Cell cell = (com.aspose.words.Cell) field.getAncestor(NodeType.CELL);

                        double maxWidth = cell.getCellFormat().getWidth()
                                - cell.getCellFormat().getLeftPadding()
                                - cell.getCellFormat().getRightPadding();

                        int lines = calculateWrappedLines(text, fontName, fontSize, maxWidth);

                        double textHeight = calculateTextHeight(lines, fontSize);

                        double allowedHeight = row.getRowFormat().getHeight();

                        if (textHeight > allowedHeight) {
                            detail.comments.append("Field: " + field.getName() + " , Data: " + field.getResult() + " , Text Height: " + textHeight + " , Allowed Height: " + allowedHeight);
                            isOverflow = true;
                            log.info(detail.comments.toString());
                            return isOverflow;
                        }
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return isOverflow;
    }

    public boolean isFormFieldOverflowDetectedOtherApproach(com.aspose.words.Document doc, FormField field) throws Exception {
        doc.updatePageLayout();
        com.aspose.words.Row row = (com.aspose.words.Row) field.getAncestor(NodeType.ROW);
        boolean isPara = false;
        boolean isOverflow = false;
        if (row != null) {
            if (row.getRowFormat().getHeightRule() == HeightRule.EXACTLY) {
                // Approach 1: Check if the form field is inside a paragraph and has an associated layout entity
                Paragraph para = (Paragraph) field.getAncestor(NodeType.PARAGRAPH);
                doc.updatePageLayout();
                LayoutCollector collector = new LayoutCollector(doc);
                LayoutEnumerator enumerator = new LayoutEnumerator(doc);
                Object entity = collector.getEntity(para);
                isPara = entity != null;

                if (isPara) {
                    enumerator.setCurrent(collector.getEntity(para));
                    double totalHeight = 0;
                    if (enumerator.moveFirstChild()) {
                        do {
                            if (enumerator.getType() == LayoutEntityType.LINE) {
                                Rectangle2D rect = enumerator.getRectangle();
                                totalHeight += rect.getHeight();
                            }
                        } while (enumerator.moveNext());
                    }
                }

                // Approach 2: Check if the form field's layout rectangle exceeds the row height
                enumerator.setCurrent(collector.getEntity(field));
                Rectangle2D rect = enumerator.getRectangle();
                double renderedHeight = rect.getHeight();
                double rowHeight = row.getRowFormat().getHeight();
                if (renderedHeight > rowHeight)
                    isOverflow = true;
            }

        }
        return isPara && isOverflow;
    }

    private double calculateTextWidth(String text, String fontName, int fontSize) {

        Font font = new Font(fontName, Font.PLAIN, fontSize);

        FontRenderContext frc = new FontRenderContext(new AffineTransform(), true, true);

        return font.getStringBounds(text, frc).getWidth();
    }

    private int calculateWrappedLines(String text,
                                      String fontName,
                                      int fontSize,
                                      double maxWidth) {

        String[] words = text.split(" ");
        int lines = 1;
        StringBuilder currentLine = new StringBuilder();

        for (String word : words) {

            String testLine = currentLine.length() == 0
                    ? word
                    : currentLine + " " + word;

            double width = calculateTextWidth(testLine, fontName, fontSize);

            if (width > maxWidth) {
                lines++;
                currentLine = new StringBuilder(word);
            } else {
                currentLine = new StringBuilder(testLine);
            }
        }

        return lines;
    }

    private double calculateTextHeight(int lineCount, int fontSize) {

        double lineHeight = fontSize * 1.2; // Word approx line spacing

        return lineCount * lineHeight;
    }


    private double getRenderedParagraphHeight(com.aspose.words.Document doc, Paragraph para) throws Exception {

        doc.updatePageLayout();
        LayoutCollector collector = new LayoutCollector(doc);
        LayoutEnumerator enumerator = new LayoutEnumerator(doc);

        Object entity = collector.getEntity(para);
        if (entity == null) return 0;

        enumerator.setCurrent(entity);

        double totalHeight = 0;

        com.aspose.words.Row row = (com.aspose.words.Row) para.getAncestor(NodeType.ROW);

        Object rowEntity = collector.getEntity(row);

        if (rowEntity != null) {
            enumerator.setCurrent(rowEntity);
            double renderedRowHeight = enumerator.getRectangle().getHeight();

            double fixedHeight = row.getRowFormat().getHeight();

            if (row.getRowFormat().getHeightRule() == HeightRule.EXACTLY
                    && renderedRowHeight > fixedHeight) {

                System.out.println("Content is trimmed in this row.");
            }
        }

        return totalHeight;
    }


    private boolean isFormField(com.aspose.pdf.Document document, Detail detail) {
        try {
            if (null != document && null != document.getForm() && null != document.getForm().getFields() && document.getForm().getFields().length > 0) {
                for (Field field : document.getForm().getFields()) {
                    if (field instanceof TextBoxField) {
                        TextBoxField text = (TextBoxField) field;
//                        boolean isScrollable = text.getMultiline() && text.getScrollable() && !text.getReadOnly() && null != text.getValue();
//                        if (isScrollable) {
//                            detail.isFormField = true;
//                            return true;
//                        }
                        boolean isScrollable = text.getMultiline() && !text.getReadOnly();
                        if (isScrollable) {
                            detail.isReadOnly = text.getReadOnly();
                            detail.isScrollable = text.getScrollable();
                            detail.hasMultiline = true;
                            return true;
                        }
                    }
                }
            }
        } catch (Exception e) {
            System.out.println("Failed to detect form fields {}" + e.getMessage() + e);
//            return detail.isFormField;
            return false;
        }
//        return detail.isFormField;
        return false;
    }

    public static boolean isXFAForm(byte[] pdfFile, Detail detail) throws IOException {
        if (null != pdfFile) {
            try (ByteArrayInputStream bi = new ByteArrayInputStream(pdfFile)) {
                com.aspose.pdf.Document pdfDocument = new com.aspose.pdf.Document(bi);
                if (null != pdfDocument.getForm() && pdfDocument.getForm().hasXfa()) {
                    detail.isXfa = true;
                    return true;
                }
//                PdfFileInfo fileInfo = new PdfFileInfo(bi);
//                String creatorName = fileInfo.getCreator();
//                if (null != creatorName && creatorName.toLowerCase().contains("adobe livecycle designer")) {
//                    return true;
//                }
//
//                if ((pdfDocument.getPermissions() & Permissions.FillForm) == Permissions.FillForm)
//                    return true;
            } catch (Exception e) {
                return false;
            }
        }
        return false;
    }

    public void emailAttachmentDetail() throws Exception {
        AmazonS3 s3Client = null;
        Workbook workbook = null;
        try {
            s3Client = AmazonS3ClientBuilder.standard()
                    .withRegion(appConfigs.getS3RegionName())
                    .withCredentials(DefaultAWSCredentialsProviderChain.getInstance())
                    .build();
            // Instantiate a new Workbook object
            workbook = new Workbook(appConfigs.getExcelPath());

//            for (int sheetNo = 0; sheetNo <= workbook.getWorksheets().getCount(); sheetNo++) {
            for (int sheetNo = 0; sheetNo <= 0; sheetNo++) {
                Worksheet worksheet = workbook.getWorksheets().get(sheetNo);
                // Get the cells from the sheet
                Cells cells = worksheet.getCells();

                // Get the maximum data row
                int maxDataRow = cells.getMaxDataRow();


                // Loop through each row
                for (int i = 1; i <= maxDataRow; i++) {
                    // Get the row
                    Row row = cells.getRow(i);

                    // Loop through each cell in the row
                    // Get the cell
                    Cell emlPathCell = row.getCellOrNull(2);
                    Cell dbAttachmentsCell = row.get(4);
                    Cell statusCell = row.get(10);

                    if (emlPathCell != null) {
                        String s3Key = emlPathCell.getStringValue();
                        if (!"COMPLETED".equalsIgnoreCase(statusCell.getStringValue()) && StringUtils.isNotBlank(s3Key)) {
                            Detail detail = new Detail();
                            try {
                                byte[] emlData;
                                if (appConfigs.isLocal()) {
                                    emlData = readFileFromLocal(s3Key, workbook);
                                } else {
                                    emlData = readFileFromS3(workbook, s3Configurations.s3Client(), s3Key);
                                }
                                if (null != emlData) {
                                    detail.attachmentsFromDb = null != dbAttachmentsCell ? dbAttachmentsCell.getStringValue() : "";
                                    detail = processPart(emlData, detail, "pdf");
                                    processDetail(detail);
                                    writeDetailToExcel(detail, row, "COMPLETED");
                                } else {
                                    writeDetailToExcel(detail, row, "File not found in S3");
                                }
                            } catch (Exception e) {
                                detail.comments.append(" [" + e.getMessage() + "] ");
                                writeDetailToExcel(detail, row, "FAILED");
                            } finally {
                                writeResult(workbook, row, detail);
                            }
                            // Save the workbook
                            log.info("row finished.....      " + i);
                            if (i % 500 == 0)
                                workbook.save(appConfigs.getExcelPath());
                            if (i % 5000 == 0)
                                return;
                        }
                    }
                }
            }
        } catch (Exception e) {
            log.error("caught Exception readExcelFileRowByRow", e);
        } finally {
            if (workbook != null) {
                workbook.save(appConfigs.getExcelPath());
                log.info("Excel updated.....      ");
            }
            if (s3Client != null) {
                s3Client.shutdown();
            }
        }
    }

    public void editablePdfAttachmentDetail() throws Exception {
        AmazonS3 s3Client = null;
        Workbook workbook = null;
        try {
            s3Client = AmazonS3ClientBuilder.standard()
                    .withRegion(appConfigs.getS3RegionName())
                    .withCredentials(DefaultAWSCredentialsProviderChain.getInstance())
                    .build();
            // Instantiate a new Workbook object
            workbook = new Workbook(appConfigs.getExcelPath());

            for (int sheetNo = 2; sheetNo <= 2; sheetNo++) {
//            for (int sheetNo = 0; sheetNo <= 0; sheetNo++) {
                Worksheet worksheet = workbook.getWorksheets().get(sheetNo);
                // Get the cells from the sheet
                Cells cells = worksheet.getCells();

                // Get the maximum data row
                int maxDataRow = cells.getMaxDataRow();


                // Loop through each row
                for (int i = 1; i <= maxDataRow; i++) {
                    // Get the row
                    Row row = cells.getRow(i);

                    // Loop through each cell in the row
                    // Get the cell
                    Cell emlPathCell = row.getCellOrNull(2);
                    Cell dbAttachmentsCell = row.get(4);
                    Cell statusCell = row.get(7);

                    if (emlPathCell != null) {
                        String s3Key = emlPathCell.getStringValue();
                        if (!"COMPLETED".equalsIgnoreCase(statusCell.getStringValue()) && StringUtils.isNotBlank(s3Key)) {
                            Detail detail = new Detail();
                            try {
                                byte[] emlData;
                                if (appConfigs.isLocal()) {
                                    emlData = readFileFromLocal(s3Key, workbook);
                                } else {
                                    emlData = readFileFromS3(workbook, s3Configurations.s3Client(), s3Key);
                                }
                                if (null != emlData) {
                                    detail.attachmentsFromDb = null != dbAttachmentsCell ? dbAttachmentsCell.getStringValue() : "";
                                    detail = processPart(emlData, detail, "pdf");
                                    writeEditablePdfDetailToExcel(detail, row, "COMPLETED");
                                } else {
                                    writeEditablePdfDetailToExcel(detail, row, "File not found in S3");
                                    log.info("row failed.....      " + i + " File not found in S3");
                                }
                            } catch (Exception e) {
                                detail.comments.append(" [" + e.getMessage() + "] ");
                                writeEditablePdfDetailToExcel(detail, row, "FAILED");
                            } finally {
                                writeResult(workbook, row, detail);
                            }
                            // Save the workbook
                            log.info("row finished.....      " + i);
                            if (i % 1000 == 0)
                                workbook.save(appConfigs.getExcelPath());
//                            if (i % 40000 == 0)
//                                return;
                        }
                    }
                }
            }
        } catch (Exception e) {
            log.error("caught Exception readExcelFileRowByRow", e);
        } finally {
            if (workbook != null) {
                workbook.save(appConfigs.getExcelPath());
                log.info("Excel updated.....      ");
            }
            if (s3Client != null) {
                s3Client.shutdown();
            }
        }
    }

    private void writeDocResult(Workbook workbook, Row row, Detail detail) {
        if (detail.hasMultiline) {

        }
    }

    private void writeResult(Workbook workbook, Row row, Detail detail) {
        if (detail.hasEditablePdf) {
            Worksheet resultSheet = workbook.getWorksheets().get(5);
            Cells resultCells = resultSheet.getCells();
            int j = resultCells.getMaxDataRow() + 1;
            Row resultRow = resultCells.getRow(j);

            Cell hasEditablePdfCell = resultRow.get(0); // editable pdf
            hasEditablePdfCell.setValue(detail.hasEditablePdf);
//
//            Cell isReadOnlyCell = resultRow.get(1); // readonly pdf
//            isReadOnlyCell.setValue(detail.isReadOnly);

//            Cell isFormFieldCell = resultRow.get(2); // multiline pdf
//            isFormFieldCell.setValue(detail.hasMultiline);

            Cell isScrollableCell = resultRow.get(1); // scrollable textfield
            isScrollableCell.setValue(detail.isScrollable);

            Cell isXfaCell = resultRow.get(2); // xfa pdf
            isXfaCell.setValue(detail.isXfa);

            Cell emailReceivedDateCell = resultRow.get(3); // Date
            emailReceivedDateCell.setValue(row.get(1).getStringValue());

            Cell receiptNumberCell = resultRow.get(4); // Receipt
            receiptNumberCell.setValue(row.get(0).getStringValue());

            Cell fileNameCell = resultRow.get(5); //editable filename
            fileNameCell.setValue(detail.fileName);

            Cell s3PathCell = resultRow.get(6);// s3 path
            s3PathCell.setValue(row.get(2).getStringValue());

            Cell messageUidCell = resultRow.get(7);// Message uid
            messageUidCell.setValue(row.get(3).getStringValue());

            Cell attachmentsFromUtilityCell = resultRow.get(8); //Utility attachment list
            attachmentsFromUtilityCell.setValue(detail.allAttachments);
        } /*else if (detail.isZip) {
            Worksheet resultSheet = workbook.getWorksheets().get(1);
            Cells resultCells = resultSheet.getCells();
            int j = resultCells.getMaxDataRow() + 1;
            Row resultRow = resultCells.getRow(j);

            Cell emailReceivedDateCell = resultRow.get(1); // Date
            emailReceivedDateCell.setValue(row.get(1));

            Cell zipFolderCell = resultRow.get(2); // folder
            zipFolderCell.setValue(detail.isZipFolder);

            Cell receiptNumberCell = resultRow.get(3); // Receipt
            receiptNumberCell.setValue(row.get(0).getStringValue());

            Cell missingCountCell = resultRow.get(4); //missing zip attachments
            missingCountCell.setValue(detail.missingFileCount);

            Cell zipEntriesCount = resultRow.get(5); //zip entries count
            zipEntriesCount.setValue(detail.zipCount);

            Cell otherAttachmentCountCell = resultRow.get(6); //other attachment count
            otherAttachmentCountCell.setValue(detail.otherAttachmentCount);

            Cell attachmentsInDbCountCell = resultRow.get(7); //DB attachment count
            attachmentsInDbCountCell.setValue(detail.dbCount);

            Cell attachmentsInDbCell = resultRow.get(8); //DB attachments list
            attachmentsInDbCell.setValue(detail.attachmentsFromDb);

            Cell attachmentsCountFromUtilityCell = resultRow.get(9); //Utility attachment list count
            attachmentsCountFromUtilityCell.setValue(detail.allAttachmentCount);

            Cell attachmentsFromUtilityCell = resultRow.get(10); //Utility attachment list
            attachmentsFromUtilityCell.setValue(detail.allAttachments);

            Cell messageUidCell = resultRow.get(11);// Message uid
            messageUidCell.setValue(row.get(3).getStringValue());

            if (detail.missingFileCount > 0) {
                int colCount = 13;
                for (int col = 0; col < colCount; col++) {
                    Cell cell = resultRow.get(col);
                    Style style = cell.getStyle();
                    style.setPattern(BackgroundType.SOLID);
                    style.setForegroundColor(Color.getIndianRed());
                    cell.setStyle(style);
                }
            } else if (detail.isZipEncrypted) {
                int colCount = 13;
                for (int col = 0; col < colCount; col++) {
                    Cell cell = resultRow.get(col);
                    Style style = cell.getStyle();
                    style.setPattern(BackgroundType.SOLID);
                    style.setForegroundColor(Color.getBlue());
                    cell.setStyle(style);
                }
            } else if (detail.isZipFolder) {
                int colCount = 13;
                for (int col = 0; col < colCount; col++) {
                    Cell cell = resultRow.get(col);
                    Style style = cell.getStyle();
                    style.setPattern(BackgroundType.SOLID);
                    style.setForegroundColor(Color.getYellowGreen());
                    cell.setStyle(style);
                }
            } else if (detail.isZip) {
                int colCount = 13;
                for (int col = 0; col < colCount; col++) {
                    Cell cell = resultRow.get(col);
                    Style style = cell.getStyle();
                    style.setPattern(BackgroundType.SOLID);
                    style.setForegroundColor(Color.getYellow());
                    cell.setStyle(style);
                }
            }
        } */
    }

    private void writeEditablePdfDetailToExcel(Detail detail, Row row, String status) {
        detail.allAttachments = detail.allAttachmentsBuilder.toString().replaceFirst(",", "");
        Cell hasEditablePdfCell = row.get(5);
        Cell allAttachmentNamesCell = row.get(6);
        Cell statusCell = row.get(7);
        Cell commentsCell = row.get(8);
        hasEditablePdfCell.setValue(String.valueOf(detail.hasEditablePdf));
        commentsCell.setValue(detail.comments.toString());
        allAttachmentNamesCell.setValue(detail.allAttachments);
        statusCell.setValue(status);
    }

    private void writeDetailToExcel(Detail detail, Row row, String status) {
        Cell isZippedCell = row.get(5);
        Cell isZippedFolderCell = row.get(6);
        Cell isZipEncryptedCell = row.get(7);
        Cell zipEntriesCell = row.get(8);
        Cell allAttachmentNamesCell = row.get(9);
//        Cell attachmentMatchedCell = row.get(10);
        Cell statusCell = row.get(10);
        Cell commentsCell = row.get(11);
        Cell zipCountCell = row.get(12);
        Cell dbCountCell = row.get(13);
        Cell missingCountCell = row.get(14);

        if (detail.isZip) {
            isZippedCell.setValue(String.valueOf(detail.isZip));
            isZippedFolderCell.setValue(String.valueOf(detail.isZipFolder));
            isZipEncryptedCell.setValue(String.valueOf(detail.isZipEncrypted));
            zipEntriesCell.setValue(detail.zipEntries);
        } else {
            isZippedCell.setValue(String.valueOf(detail.isZip));
        }
        commentsCell.setValue(detail.comments.toString());
//        attachmentMatchedCell.setValue(String.valueOf(detail.attachmentsMatched));
        allAttachmentNamesCell.setValue(detail.allAttachments);
        dbCountCell.setValue(detail.dbCount);
        zipCountCell.setValue(detail.zipCount);
        missingCountCell.setValue(detail.missingFileCount);
        statusCell.setValue(status);
    }

    private void processDetail(Detail detail) {
        detail.allAttachments = detail.allAttachmentsBuilder.toString().replaceFirst(",", "");
        detail.zipEntries = detail.zipEntriesBuilder.toString().replaceFirst(",", "");

        ArrayList<String> zipList = new ArrayList<>();
        Arrays.stream(detail.zipEntries.split("\\s*,\\s*"))
                .map(String::trim)
                .filter(s -> !s.isEmpty())
                .forEach(zipList::add);
        ArrayList<String> otherAttachmentList = new ArrayList<>();
        Arrays.stream(detail.allAttachments.split("\\s*,\\s*"))
                .map(String::trim)
                .filter(file -> !file.toLowerCase().endsWith(".zip"))
                .forEach(otherAttachmentList::add);
        ArrayList<String> dbList = new ArrayList<>();
        if (detail.isZip) {
            Pattern pattern = Pattern.compile("[^\\s]+?\\.(pdf|xls|xlsx|csv|doc|docx|ppt|pptx|txt|rtf|jpg|jpeg|png|gif|bmp|tiff|zip|rar|7z|tar|gz|xml|msg|eml|rpmsg|html)", Pattern.CASE_INSENSITIVE);
            Matcher matcher = pattern.matcher(detail.attachmentsFromDb);
            while (matcher.find()) {
                dbList.add(matcher.group());
            }
        } else {
            Arrays.stream(detail.attachmentsFromDb.split("\\s*,\\s*"))
                    .map(String::trim)
                    .filter(s -> !s.isEmpty())
                    .forEach(dbList::add);
        }


        // Find counts
        int processedCount = zipList.size() + otherAttachmentList.size();

        detail.missingFileCount = processedCount - dbList.size();
        detail.dbCount = dbList.size();
        detail.zipCount = zipList.size();
        detail.otherAttachmentCount = otherAttachmentList.size();
        detail.allAttachmentCount = (int) Arrays.stream(detail.allAttachments.split("\\s*,\\s*")).map(String::trim).count();
    }

    private Detail processPart(byte[] emlData, Detail detail, String fileTypeToCheck) {
        Message message = null;
        try {
            message = new MimeMessage(null, new ByteArrayInputStream(emlData));

            Object content = message.getContent();
            if (content instanceof Multipart) {
                Multipart multipart = (Multipart) content;
                for (int j = 0; j < multipart.getCount(); j++) {
                    BodyPart bodyPart = multipart.getBodyPart(j);
                    if (isAttachment(bodyPart)) {
                        String fileName = bodyPart.getFileName();
                        String fileType = extractFileType(fileName);
                        if ("zip".equalsIgnoreCase(fileType)) {
//                            detail.isZip = true;
                            processZip(fileName, IOUtils.toByteArray(bodyPart.getInputStream()), detail);
                        } else if ("eml".equalsIgnoreCase(fileType) || "msg".equalsIgnoreCase(fileType) || "rpmsg".equalsIgnoreCase(fileType)) {
                            processPart(IOUtils.toByteArray(bodyPart.getInputStream()), detail, fileTypeToCheck);
                        }
                        if ("pdf".equalsIgnoreCase(fileType) && "pdf".equalsIgnoreCase(fileTypeToCheck)) {
                            processPdfAttachment(fileName, IOUtils.toByteArray(bodyPart.getInputStream()), detail);
                        } else if (("docx".equalsIgnoreCase(fileType) || "doc".equalsIgnoreCase(fileType)) && "word".equalsIgnoreCase(fileTypeToCheck)) {
                            processWordAttachment(fileName, IOUtils.toByteArray(bodyPart.getInputStream()), detail);
                        }
                        detail.allAttachmentsBuilder.append(",");
                        detail.allAttachmentsBuilder.append(fileName);
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return detail;
    }

    private void processPdfAttachment(String fileName, byte[] contents, Detail detail) throws IOException {
        com.aspose.pdf.Document document = new com.aspose.pdf.Document(contents);
        if (isFormField(document, detail) || isXFAForm(contents, detail)) {
            detail.comments.append(" [Editable PDF Form Detected]");
            detail.hasEditablePdf = true;
            detail.fileName = fileName;
        }
    }

    public void processWordAttachment(String fileName, byte[] contents, Detail detail) throws IOException {
        Path tempFile = null;
        try {
            tempFile = Paths.get(appConfigs.getTempPath(), fileName);
            Files.createFile(tempFile);
            Files.write(tempFile, contents);
            com.aspose.words.Document document = new com.aspose.words.Document(tempFile.toAbsolutePath().toString());
            if (null == detail)
                detail = new Detail();
            if (isMultilineWordDoc(document, detail)) {
                detail.hasMultiline = true;
                detail.fileName = fileName;
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            Files.deleteIfExists(tempFile);
        }
    }

    public void processWordAttachmentTest(String fileName, Detail detail) throws IOException {
        try {
            com.aspose.words.Document document = new com.aspose.words.Document(fileName);
            if (null == detail)
                detail = new Detail();
            if (isMultilineWordDoc(document, detail)) {
                detail.hasMultiline = true;
                detail.fileName = fileName;
            }
            System.out.println("Processed Word document: " + fileName + ", hasMultiline: " + detail.hasMultiline + " [" + detail.comments.toString() + "]");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private Run getFirstResultRun(FormField ff) {

        Node current = ff.getNextSibling();

        // Move until separator
        while (current != null && current.getNodeType() != NodeType.FIELD_SEPARATOR) {
            current = current.getNextSibling();
        }

        if (current == null) return null;

        // Move to first result run
        current = current.getNextSibling();

        while (current != null && current.getNodeType() != NodeType.FIELD_END) {

            if (current.getNodeType() == NodeType.RUN) {
                return (Run) current;
            }

            current = current.getNextSibling();
        }

        return null;
    }


    class Detail {
        boolean isZip;
        boolean isZipFolder;
        StringBuilder allAttachmentsBuilder = new StringBuilder();
        StringBuilder zipEntriesBuilder = new StringBuilder();
        String allAttachments = "";
        String zipEntries = "";
        String attachmentsFromDb = "";

        boolean isZipEncrypted;
        int missingFileCount;
        int dbCount;
        int zipCount;
        int otherAttachmentCount;
        int allAttachmentCount;
        StringBuilder comments = new StringBuilder();

        boolean hasEditablePdf;
        boolean isXfa;
        boolean isFormField;
        boolean isReadOnly;
        boolean hasMultiline;
        boolean isScrollable;
        String fileName;
    }

    private Detail processFile(byte[] emlData, Detail detail, String pdfFileName, String receiptNumber) {
        Message message = null;
        try {
            message = new MimeMessage(null, new ByteArrayInputStream(emlData));

            Object content = message.getContent();
            if (content instanceof Multipart) {
                Multipart multipart = (Multipart) content;
                for (int j = 0; j < multipart.getCount(); j++) {
                    BodyPart bodyPart = multipart.getBodyPart(j);
                    if (isAttachment(bodyPart)) {
                        String fileName = bodyPart.getFileName();
                        String fileType = extractFileType(fileName);
                        if ("zip".equalsIgnoreCase(fileType)) {
//                            detail.isZip = true;
                            processZip(fileName, IOUtils.toByteArray(bodyPart.getInputStream()), detail);
                        } else if ("eml".equalsIgnoreCase(fileType) || "msg".equalsIgnoreCase(fileType) || "rpmsg".equalsIgnoreCase(fileType)) {
                            processFile(IOUtils.toByteArray(bodyPart.getInputStream()), detail, fileName, receiptNumber);
                        }
                        if (pdfFileName.equals(fileName)) {
                            if (!"pdf".equalsIgnoreCase(fileType)) {
                                convertAndSaveFile(fileName, IOUtils.toByteArray(bodyPart.getInputStream()), receiptNumber);
                            } else {
                                flattenAndSaveFile(fileName, IOUtils.toByteArray(bodyPart.getInputStream()), receiptNumber);
                            }
                        }
                        detail.allAttachmentsBuilder.append(",");
                        detail.allAttachmentsBuilder.append(fileName);
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return detail;
    }


    private byte[] convertToPdf(byte[] contents) {
        try (ByteArrayInputStream bi = new ByteArrayInputStream(contents);
             ByteArrayOutputStream os = new ByteArrayOutputStream()) {
            com.aspose.words.Document doc = new com.aspose.words.Document(bi);
            com.aspose.words.PdfSaveOptions saveOptions = new com.aspose.words.PdfSaveOptions();
            saveOptions.setSaveFormat(SaveFormat.PDF);
            saveOptions.setMemoryOptimization(true);
            wordDocumentExtraction(doc);
            doc.save(os, saveOptions);
            return os.toByteArray();
        } catch (Exception e) {
            System.out.println(e);
        }
        return null;
    }

    private static void wordDocumentExtraction(com.aspose.words.Document doc) throws Exception {
        try {
            if (null != doc && null != doc.getSections()) {
                for (Section section : doc.getSections()) {
                    if (null != section.getBody() && null != section.getBody().getTables() && section.getBody().getTables().getCount() > 0) {
                        for (Table table : section.getBody().getTables()) {
                            if (table.getTextWrapping() == TextWrapping.AROUND
                                    && table.getAllowOverlap()
                                    && table.getRelativeHorizontalAlignment() != HorizontalAlignment.RIGHT) {
                                table.setTextWrapping(TextWrapping.NONE);
                            }
                        }
                    }
                }
            }
        } catch (Exception e) {
            System.out.println("getting Exception from @wordDocumentExtraction ::{}" + e);
        }
    }

    private void convertAndSaveFile(String fileName, byte[] contents, String receiptNumber) {
        FileOutputStream fos = null;
        try {
            // Create receipt directory if not exists
            File receiptDir = new File(appConfigs.getOutputPath(), receiptNumber);
            if (!receiptDir.exists()) {
                receiptDir.mkdirs();
            }

            // 1. Save original PDF
            File originalFile = new File(receiptDir, "ORIGINAL_" + fileName);
            fos = new FileOutputStream(originalFile);
            fos.write(contents);
            fos.flush();
            fos.close();

            // 2. Load PDF from bytes and flatten
            byte[] convertedPdfContent = convertToPdf(contents);

            // 3. Save converted PDF
            if (convertedPdfContent != null) {
                File convertedFile = new File(receiptDir, "CONVERTED_" + FilenameUtils.getBaseName(fileName) + ".pdf");
                FileUtils.writeByteArrayToFile(convertedFile, convertedPdfContent);
                log.info("Saved original and Converted PDF for receiptNumber={}", receiptNumber);
            } else {
                log.info("Fail to Converted PDF for receiptNumber={}", receiptNumber);
            }

        } catch (Exception e) {
            log.error("Error saving PDFs for receipt number: {}", receiptNumber, e);
        } finally {
            if (fos != null) {
                try {
                    fos.close();
                } catch (IOException ignored) {
                }
            }
        }
    }

    private void flattenAndSaveFile(String fileName, byte[] contents, String receiptNumber) {
        FileOutputStream fos = null;
        try {
            // Create receipt directory if not exists
            File receiptDir = new File(appConfigs.getOutputPath(), receiptNumber);
            if (!receiptDir.exists()) {
                receiptDir.mkdirs();
            }

            // 1. Save original PDF
            File originalFile = new File(receiptDir, "ORIGINAL_" + fileName);
            fos = new FileOutputStream(originalFile);
            fos.write(contents);
            fos.flush();
            fos.close();

            // 2. Load PDF from bytes and flatten
            com.aspose.pdf.Document document = new com.aspose.pdf.Document(contents);
            document.flatten();

            // 3. Save flattened PDF
            File flattenedFile = new File(receiptDir, "FLATTENED_" + fileName);
            document.save(flattenedFile.getAbsolutePath());

            log.info("Saved original and flattened PDF for receiptNumber={}", receiptNumber);

        } catch (Exception e) {
            log.error("Error saving PDFs for receipt number: {}", receiptNumber, e);
        } finally {
            if (fos != null) {
                try {
                    fos.close();
                } catch (IOException ignored) {
                }
            }
        }
    }


    private void processZip(String fileName, byte[] zipContent, Detail detail) throws IOException {
        File tempZipFile = File.createTempFile(appConfigs.getTempPath() + fileName, "");
        try {
            FileOutputStream fos = new FileOutputStream(tempZipFile);
            fos.write(zipContent);
            ZipFile zipFile = new ZipFile(tempZipFile);
            if (zipFile.isEncrypted()) {
                detail.isZipEncrypted = true;
                detail.comments.append(" [Encrypted] ");
                return;
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (null != tempZipFile && tempZipFile.exists())
                tempZipFile.delete();
        }
        try (ByteArrayInputStream bi = new ByteArrayInputStream(zipContent);
             ZipInputStream zis = new ZipInputStream(bi)) {
            ZipEntry entry;
            int n;
            byte[] buf = new byte[1024];
            while ((entry = zis.getNextEntry()) != null) {
                // Uncomment to get zip entries details
              /*  String entryName = entry.getName();
                detail.zipEntriesBuilder.append(",");
                detail.zipEntriesBuilder.append(entryName);
                if (entry.isDirectory()) {
                    detail.isZipFolder = true;
                }
                if (entryName.contains("/") && !entryName.startsWith("/")) {
                    detail.isZipFolder = true;
                }*/
                // Process only PDF files
                String entryFileName = MimeUtility.decodeText(Objects.requireNonNull(entry.getName()));
                String fileType = FilenameUtils.getExtension(entryFileName);
                if ("pdf".equalsIgnoreCase(fileType)) {
                    try (ByteArrayOutputStream baop = new ByteArrayOutputStream()) {
                        while ((n = zis.read(buf, 0, 1024)) > -1) {
                            baop.write(buf, 0, n);
                        }
                        processPdfAttachment(entryFileName, baop.toByteArray(), detail);
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public boolean isAttachment(Part part) {
        try {
            String disposition = part.getDisposition();
            return part.getFileName() != null
                    && (disposition == null || Part.ATTACHMENT.equalsIgnoreCase(disposition) || Part.INLINE.equalsIgnoreCase(disposition));
        } catch (Exception e) {
            return false;
        }
    }

    public String extractFileType(String fileName) {
        if (StringUtils.isNotBlank(fileName)) {
            String fileType = FilenameUtils.getExtension(fileName);
            if (StringUtils.isNotBlank(fileType)) {
                if (fileType.contains("pdf") || fileType.contains("PDF")) return "pdf";
                if (fileType.contains("docx") || fileType.contains("DOCX")) return "docx";
                if (fileType.contains("doc") || fileType.contains("DOC")) return "doc";
                if (fileType.contains("xlsx") || fileType.contains("XLSX")) return "xlsx";
                if (fileType.contains("xls") || fileType.contains("XLS")) return "xls";
                if (fileType.contains("zip") || fileType.contains("ZIP")) return "zip";
            }
            return fileType;
        }
        return "";
    }


    public void readExcelFileRowByRow() throws Exception {
        AmazonS3 s3Client = null;
        Workbook workbook = null;
        try {
           /* s3Client = AmazonS3Client.builder()
                    .withRegion(appConfigs.getS3RegionName())
                    .withCredentials(DefaultAWSCredentialsProviderChain.getInstance())
                    .build();*/
            s3Client = AmazonS3ClientBuilder.standard()
                    .withRegion(appConfigs.getS3RegionName())
                    .withCredentials(DefaultAWSCredentialsProviderChain.getInstance())
                    .build();
            //s3Client = AmazonS3ClientBuilder.defaultClient();
            // Instantiate a new Workbook object
            workbook = new Workbook(appConfigs.getExcelPath());

            // Get the first worksheet
            Worksheet worksheet = workbook.getWorksheets().get(0);

            // Get the cells from the sheet
            Cells cells = worksheet.getCells();

            // Get the maximum data row
            int maxDataRow = cells.getMaxDataRow();

            // Loop through each row
            for (int i = 1; i <= maxDataRow; i++) {
                // Get the row
                Row row = cells.getRow(i);

                // Loop through each cell in the row
                // Get the cell
                Cell cell = row.getCellOrNull(6);
                Cell statusCell = row.getCellOrNull(7);

                // If the cell is not null, print its text
                if (cell != null && statusCell == null) {
                    String s3Key = cell.getStringValue();
                    if (s3Key != null && !s3Key.isEmpty()) {
                        com.aspose.pdf.Document pdfDocument = null;
                        if (isLocal) {
                            pdfDocument = new Document(row.getCellOrNull(6).getStringValue());
                        } else {
                            try {
                                byte[] pdfBytes = readFileFromS3(workbook, s3Client, row.getCellOrNull(6).getStringValue());
                                pdfDocument = new Document(new ByteArrayInputStream(pdfBytes));
                            } catch (Exception e) {
                                Cell processStatusCell = cells.get(i, 7);
                                processStatusCell.setValue("F");
                                Cell errorLogCell = cells.get(i, 8);
                                errorLogCell.setValue(e.getMessage());
                                log.error("caught Exception with pdfBytes" + row.getCellOrNull(6).getStringValue(), e);
                                continue;
                            }
                        }

                        if (pdfDocument != null) {
                            String receiptNo = row.getCellOrNull(0).getStringValue();
                            String docId = row.getCellOrNull(1).getStringValue();
                            try {
                                boolean anyFileCorrupted = determineFileTypeAndSaveFile(pdfDocument, receiptNo, docId);
                                Cell processStatusCell = cells.get(i, 7);
                                processStatusCell.setValue("S");
                                if (anyFileCorrupted) {
                                    Cell impactedFileCell = cells.get(i, 9);
                                    impactedFileCell.setValue("Yes");
                                }
                            } catch (Exception e) {
                                Cell processStatusCell = cells.get(i, 7);
                                processStatusCell.setValue("F");
                                Cell errorLogCell = cells.get(i, 8);
                                errorLogCell.setValue(e.getMessage());
                                log.error("Error in determineFileTypeAndSaveFile receiptNo for " + receiptNo + " ,for docId " + docId, e);
                            }
                        } else {
                            Cell processStatusCell = cells.get(i, 7);
                            processStatusCell.setValue("F");
                            Cell errorLogCell = cells.get(i, 8);
                            errorLogCell.setValue("File Not Found in S3");
                        }
                        // Save the workbook
                        log.info("row finished.....      " + i);
                        if (i % 50 == 0)
                            workbook.save(appConfigs.getExcelPath());
                    }
                }
            }
            workbook.save(appConfigs.getExcelPath());
            log.info("Excel updated.....      ");
        } catch (Exception e) {
            if (workbook != null) {
                workbook.save(appConfigs.getExcelPath());
            }
            log.error("caught Exception readExcelFileRowByRow", e);
        } finally {
            if (s3Client != null) {
                s3Client.shutdown();
            }
        }
    }

    public void readExcelFileRowByRowForOriginal() throws Exception {
        AmazonS3 s3Client = null;
        Workbook workbook = null;
        try {
           /* s3Client = AmazonS3Client.builder()
                    .withRegion(appConfigs.getS3RegionName())
                    .withCredentials(DefaultAWSCredentialsProviderChain.getInstance())
                    .build();*/
            s3Client = AmazonS3ClientBuilder.standard()
                    .withRegion(appConfigs.getS3RegionName())
                    .withCredentials(DefaultAWSCredentialsProviderChain.getInstance())
                    .build();
            //s3Client = AmazonS3ClientBuilder.defaultClient();
            // Instantiate a new Workbook object
            workbook = new Workbook(appConfigs.getExcelPath());

            // Get the first worksheet
            Worksheet worksheet = workbook.getWorksheets().get(0);

            // Get the cells from the sheet
            Cells cells = worksheet.getCells();

            // Get the maximum data row
            int maxDataRow = cells.getMaxDataRow();

            // Loop through each row
            for (int i = 0; i <= maxDataRow; i++) {
                // Get the row
                Row row = cells.getRow(i);

                // Loop through each cell in the row
                // Get the cell
                Cell cell = row.getCellOrNull(6);
                Cell statusCell = row.getCellOrNull(7);

                // If the cell is not null, print its text
                if (cell != null && (statusCell == null || "".equals(statusCell.getStringValue()))) {
                    String s3Key = cell.getStringValue();
                    if (s3Key != null && !s3Key.isEmpty()) {
                        com.aspose.pdf.Document pdfDocument = null;
                        try {
                            byte[] pdfBytes = readFileFromS3(workbook, s3Client, row.getCellOrNull(6).getStringValue());
                            if (pdfBytes != null) {
                                pdfDocument = new Document(new ByteArrayInputStream(pdfBytes));
                                String receiptNo = row.getCellOrNull(0).getStringValue();
                                String docId = row.getCellOrNull(1).getStringValue();
                                StringBuilder affectedFileNames = new StringBuilder();
                                for (FileSpecification fileSpec : pdfDocument.getEmbeddedFiles()) {
                                    String embeddedFileName = fileSpec.getName();
                                    boolean anyEmbeddedFileCorrupted = isEmbeddedFileCorrupted(embeddedFileName);
                                    if (anyEmbeddedFileCorrupted && embeddedFileName.length() >= 60) {
                                        Cell impactedCell = cells.get(i, 9);
                                        impactedCell.setValue("Yes");
                                        if (!affectedFileNames.isEmpty()) {
                                            affectedFileNames.append(",").append(embeddedFileName);
                                        } else {
                                            affectedFileNames.append(embeddedFileName);
                                        }
                                    }
                                }
                                Cell impactedFiles = cells.get(i, 10);
                                impactedFiles.setValue(affectedFileNames.toString());
                                pdfDocument.save("C:\\BMS_MAGS-1039546\\Output\\OriginalRCTMsgBodyFiles\\" + receiptNo + "_" + docId + "_original.pdf");
                                Cell processStatusCell = cells.get(i, 7);
                                processStatusCell.setValue("S");
                            } else {
                                Cell processStatusCell = cells.get(i, 7);
                                processStatusCell.setValue("F");
                                Cell errorLogCell = cells.get(i, 8);
                                errorLogCell.setValue("Source PDF is null or corrupted");
                            }
                        } catch (Exception e) {
                            Cell processStatusCell = cells.get(i, 7);
                            processStatusCell.setValue("F");
                            Cell errorLogCell = cells.get(i, 8);
                            errorLogCell.setValue(e.getMessage());
                            log.error("caught Exception with pdfBytes" + row.getCellOrNull(6).getStringValue(), e);
                            continue;
                        }

                        // Save the workbook
                        log.info("row finished.....      " + i);
                        if (i % 50 == 0)
                            workbook.save(appConfigs.getExcelPath());
                    }
                }
            }
            workbook.save(appConfigs.getExcelPath());
            log.info("Excel updated.....      ");
        } catch (Exception e) {
            if (workbook != null) {
                workbook.save(appConfigs.getExcelPath());
            }
            log.error("caught Exception readExcelFileRowByRow", e);
        } finally {
            if (s3Client != null) {
                s3Client.shutdown();
            }
        }
    }

    public void readExcelFileRowByRowCopyImpacted() throws Exception {
        //AmazonS3 s3Client = null;
        Workbook workbook = null;
        try {
           /* s3Client = AmazonS3Client.builder()
                    .withRegion(appConfigs.getS3RegionName())
                    .withCredentials(DefaultAWSCredentialsProviderChain.getInstance())
                    .build();*/
           /* s3Client = AmazonS3ClientBuilder.standard()
                    .withRegion(appConfigs.getS3RegionName())
                    .withCredentials(DefaultAWSCredentialsProviderChain.getInstance())
                    .build();*/
            //s3Client = AmazonS3ClientBuilder.defaultClient();
            // Instantiate a new Workbook object
            workbook = new Workbook(appConfigs.getExcelPath());

            // Get the first worksheet
            Worksheet worksheet = workbook.getWorksheets().get(0);

            // Get the cells from the sheet
            Cells cells = worksheet.getCells();

            // Get the maximum data row
            int maxDataRow = cells.getMaxDataRow();

            // Loop through each row
            for (int i = 0; i <= maxDataRow; i++) {
                // Get the row
                Row row = cells.getRow(i);

                // Loop through each cell in the row
                // Get the cell
                Cell actualImpacted = row.getCellOrNull(10);
                Cell statusCell = row.getCellOrNull(7);

                // If the cell is not null, print its text
                if ((actualImpacted != null && "Yes".equals(actualImpacted.getStringValue())) && (statusCell != null && "S".equals(statusCell.getStringValue()))) {
                    try {
                        String receiptNo = row.getCellOrNull(0).getStringValue();
                        String docId = row.getCellOrNull(1).getStringValue();
                        String srcTmpPath = "C:\\BMS_MAGS-1039546\\Output\\UpdatedRCTMsgBodyFiles\\";
                        String destTmpPath = "C:\\BMS_MAGS-1039546\\Output\\ActImpactedUpdMsgBodys";
                        String srcFileName = srcTmpPath + receiptNo + "_" + docId + "_updated.pdf";
                        copyFile(srcFileName, destTmpPath);
                        Cell copiedStatusCell = cells.get(i, 11);
                        copiedStatusCell.setValue("Copied");
                    } catch (Exception e) {
                        Cell copiedStatusCell = cells.get(i, 11);
                        copiedStatusCell.setValue("Not Copied");
                        Cell errorLogCell = cells.get(i, 12);
                        errorLogCell.setValue(e.getMessage());
                        log.error("caught Exception with pdfBytes" + row.getCellOrNull(6).getStringValue(), e);
                        continue;
                    }

                    // Save the workbook
                    log.info("row finished.....      " + i);
                    if (i % 50 == 0)
                        workbook.save(appConfigs.getExcelPath());

                }
            }
            workbook.save(appConfigs.getExcelPath());
            log.info("Excel updated.....      ");
        } catch (Exception e) {
            if (workbook != null) {
                workbook.save(appConfigs.getExcelPath());
            }
            log.error("caught Exception readExcelFileRowByRow", e);
        }
    }

    public void readExcelFileRowByRowCopyWithoutWorldWideBMS() throws Exception {
        //AmazonS3 s3Client = null;
        Workbook workbook = null;
        try {
           /* s3Client = AmazonS3Client.builder()
                    .withRegion(appConfigs.getS3RegionName())
                    .withCredentials(DefaultAWSCredentialsProviderChain.getInstance())
                    .build();*/
           /* s3Client = AmazonS3ClientBuilder.standard()
                    .withRegion(appConfigs.getS3RegionName())
                    .withCredentials(DefaultAWSCredentialsProviderChain.getInstance())
                    .build();*/
            //s3Client = AmazonS3ClientBuilder.defaultClient();
            // Instantiate a new Workbook object
            workbook = new Workbook(appConfigs.getExcelPath());

            // Get the first worksheet
            Worksheet worksheet = workbook.getWorksheets().get(0);

            // Get the cells from the sheet
            Cells cells = worksheet.getCells();

            // Get the maximum data row
            int maxDataRow = cells.getMaxDataRow();

            // Loop through each row
            for (int i = 0; i <= maxDataRow; i++) {
                // Get the row
                Row row = cells.getRow(i);

                // Loop through each cell in the row
                // Get the cell
                Cell impacted = row.getCellOrNull(9);
                Cell statusCell = row.getCellOrNull(7);

                // If the cell is not null, print its text
                if ((impacted != null && "Yes".equals(impacted.getStringValue())) && (statusCell != null && "S".equals(statusCell.getStringValue()))) {
                    try {
                        String receiptNo = row.getCellOrNull(0).getStringValue();
                        String docId = row.getCellOrNull(1).getStringValue();
                        String toMailId = row.getCellOrNull(4) != null ? row.getCellOrNull(4).getStringValue() : null;
                        String ccMailId = row.getCellOrNull(11) != null ? row.getCellOrNull(11).getStringValue() : null;
                        boolean isWWSafetyBMSPresent = false;
                        if (toMailId != null && !toMailId.isEmpty()) {
                            List<String> toMailIds = Arrays.asList(toMailId.split(";"));
                            Optional isWWSBms = toMailIds.stream().filter(mailId -> mailId.equalsIgnoreCase("WORLDWIDE.SAFETY@BMS.COM")).findAny();
                            isWWSafetyBMSPresent = isWWSBms.isPresent();
                        }
                        if (!isWWSafetyBMSPresent && ccMailId != null && !ccMailId.isEmpty()) {
                            List<String> ccMailIds = Arrays.asList(ccMailId.split(";"));
                            Optional isWWSBms = ccMailIds.stream().filter(mailId -> mailId.equalsIgnoreCase("WORLDWIDE.SAFETY@BMS.COM")).findAny();
                            isWWSafetyBMSPresent = isWWSBms.isPresent();
                        }
                        if (!isWWSafetyBMSPresent) {
                            String srcTmpPath = "C:\\BMS_MAGS-1039546\\Output\\UpdatedRCTMsgBodyFiles\\";
                            String destTmpPath = "C:\\BMS_MAGS-1039546\\Output\\WithoutWWBMS";
                            String srcFileName = srcTmpPath + receiptNo + "_" + docId + "_updated.pdf";
                            copyFile(srcFileName, destTmpPath);
                            Cell copiedStatusCell = cells.get(i, 12);
                            copiedStatusCell.setValue("Copied");
                        }
                    } catch (Exception e) {
                        Cell copiedStatusCell = cells.get(i, 12);
                        copiedStatusCell.setValue("Not Copied");
                        Cell errorLogCell = cells.get(i, 13);
                        errorLogCell.setValue(e.getMessage());
                        log.error("caught Exception with pdfBytes" + row.getCellOrNull(6).getStringValue(), e);
                        continue;
                    }

                    // Save the workbook
                    log.info("row finished.....      " + i);
                    if (i % 50 == 0)
                        workbook.save(appConfigs.getExcelPath());

                }
            }
            workbook.save(appConfigs.getExcelPath());
            log.info("Excel updated.....      ");
        } catch (Exception e) {
            if (workbook != null) {
                workbook.save(appConfigs.getExcelPath());
            }
            log.error("caught Exception readExcelFileRowByRow", e);
        }
    }

    private byte[] readFileFromS3(Workbook workbook, AmazonS3 s3Client, String key) throws Exception {
        try {
            S3Object s3object = s3Client.getObject(appConfigs.getS3BucketName(), key);
            S3ObjectInputStream inputStream = s3object.getObjectContent();
            byte[] content = IOUtils.toByteArray(inputStream);
//            saveEmlToLocal(key, content);
            return content;
        } catch (Exception e) {
            log.error("caught Exception readFileFromS3", e);
            if (workbook != null) {
                workbook.save(appConfigs.getExcelPath());
                log.info("Excel updated.....      ");
            }
            System.exit(0);
        }
        return null;
    }

    private void saveEmlToLocal(String key, byte[] content) {
        try {
            String fileName = key.substring(key.lastIndexOf('/') + 1);
            String filePath = appConfigs.getEmlPath() + "/" + fileName;
            FileOutputStream os = new FileOutputStream(filePath);
            os.write(content);
            os.close();
        } catch (Exception e) {
            log.error("caught Exception saveEmlToLocal", e);
        }
    }

    private byte[] readFileFromLocal(String key, Workbook workbook) throws Exception {
        try {
            String fileName = key.substring(key.lastIndexOf('/') + 1);
            String filePath = appConfigs.getEmlPath() + "/" + fileName;
            return new FileInputStream(filePath).readAllBytes();
        } catch (Exception e) {
//            log.error("caught Exception readFileFromLocal", e);
//            if (workbook != null) {
//                workbook.save(appConfigs.getExcelPath());
//                log.info("Excel updated.....      ");
//            }
//            System.exit(0);
            return readFileFromS3(workbook, s3Configurations.s3Client(), key);
        }
//        return null;
    }

    private boolean determineFileTypeAndSaveFile(com.aspose.pdf.Document pdfDocument, String rctNo, String docId) throws
            Exception {
        boolean anyEmbeddedFileCorrupted = false;
        try {
            String tmpPath = "C:\\BMS_MAGS-1039546\\Output";
            Map<String, String> mimeExtnMap = getMimeExtnMap();

            for (FileSpecification fileSpecification : pdfDocument.getEmbeddedFiles()) {
                String embeddedFileName = fileSpecification.getName();
                anyEmbeddedFileCorrupted = isEmbeddedFileCorrupted(embeddedFileName);
                if (anyEmbeddedFileCorrupted)
                    break;
            }
            if (anyEmbeddedFileCorrupted) {
                for (FileSpecification fileSpecification : pdfDocument.getEmbeddedFiles()) {
                    System.out.println("fileSpecification.getName() = " + fileSpecification.getName());
                    String embeddedFileName = fileSpecification.getName();
                    boolean isFileCorrupted = isEmbeddedFileCorrupted(embeddedFileName);
                    if (isFileCorrupted) {
                        writeToFolder(fileSpecification, tmpPath + File.separator + "OriginalEmbeddedFiles" + File.separator + rctNo + "_" + docId + "_original_" + embeddedFileName);
                        String fileToDelete = fileSpecification.getName();
                        Tika tika = new Tika();
                        org.apache.tika.metadata.Metadata metadata = new Metadata();
                        metadata.set("resourceName", java.net.URLDecoder.decode(embeddedFileName, StandardCharsets.UTF_8));
                        BufferedInputStream bis = new BufferedInputStream(fileSpecification.getContents());
                        String mimeTypeFromFileContent = tika.getDetector().detect(bis, metadata).toString();
                        if (mimeExtnMap.get(mimeTypeFromFileContent) != null) {
                            embeddedFileName = embeddedFileName + "." + mimeExtnMap.get(mimeTypeFromFileContent);
                            writeToFolder(fileSpecification, tmpPath + File.separator + "UpdatedEmbeddedFiles" + File.separator + rctNo + "_" + docId + "_updated_" + embeddedFileName);
                            pdfDocument.getEmbeddedFiles().delete(fileToDelete);
                            pdfDocument.getEmbeddedFiles().add(new FileSpecification(fileSpecification.getContents(), embeddedFileName));
                        } else {
                            throw new Exception("Mime Type Extn not Found" + mimeTypeFromFileContent);
                        }
                    } else {
                        writeToFolder(fileSpecification, tmpPath + File.separator + "OriginalEmbeddedFiles" + File.separator + rctNo + "_" + docId + "_original_" + embeddedFileName);
                    }

                }
                pdfDocument.save(tmpPath + File.separator + "UpdatedRCTMsgBodyFiles" + File.separator + rctNo + "_" + docId + "_updated.pdf");
            }

        } catch (Exception e) {
            log.error("Exception in determineFileTypeAndSaveFile " + rctNo + "docId" + docId, e);
            throw e;
        }
        return anyEmbeddedFileCorrupted;
    }

    private static boolean isEmbeddedFileCorrupted(String embeddedFileName) {
        boolean embeddedFileCorrupted = false;
        String[] fileTypes = {"pdf", "docx", "xlsx", "pptx", "txt", "csv", "jpg", "png", "gif", "bmp", "tiff", "xls", "doc", "ppt", "html", "xml", "msg", "eml", "HEIC", "rtf", "zip", "vcf", "p7s", "mp4", "MOV", "rpmsg", "switch", "pages"};
        int index = embeddedFileName.lastIndexOf(".");
        if (index == -1) {
            embeddedFileCorrupted = true;
        } else {
            String fileType = embeddedFileName.substring(index + 1);
            if (!Arrays.asList(fileTypes).contains(fileType)) {
                embeddedFileCorrupted = true;
            }
        }
        return embeddedFileCorrupted;
    }

    private static void writeToFolder(FileSpecification fileSpecification, String fileName) throws IOException {
        try (InputStream input = fileSpecification.getContents();
             FileOutputStream output = new FileOutputStream(fileName, true);) {
            byte[] buffer = new byte[4096];
            int n = 0;
            while (-1 != (n = input.read(buffer)))
                output.write(buffer, 0, n);
            // Close InputStream object
            output.close();
            input.close();
        } catch (Exception e) {
            log.error("Exception while writing file to tmpPath rctNo " + fileName, e);
            throw e;
        } finally {
//            fileSpecification.dispose();
        }
    }


    private static Map<String, String> getMimeExtnMap() {
        Map<String, String> mimeExtnMap = new HashMap<>();
        mimeExtnMap.put("application/pdf", "pdf");
        mimeExtnMap.put("image/bmp", "bmp");
        mimeExtnMap.put("text/csv", "csv");
        mimeExtnMap.put("application/msword", "doc");
        mimeExtnMap.put("application/vnd.openxmlformats-officedocument.wordprocessingml.document", "docx");
        mimeExtnMap.put("image/gif", "gif");
        mimeExtnMap.put("image/jpeg", "jpg");
        mimeExtnMap.put("image/png", "png");
        mimeExtnMap.put("application/vnd.ms-powerpoint", "ppt");
        mimeExtnMap.put("application/vnd.openxmlformats-officedocument.presentationml.presentation", "pptx");
        mimeExtnMap.put("image/tiff", "tiff");
        mimeExtnMap.put("application/vnd.ms-excel", "xls");
        mimeExtnMap.put("text/plain", "txt");
        mimeExtnMap.put("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "xlsx");
        mimeExtnMap.put("application/xml", "xml");
        mimeExtnMap.put("application/rtf", "rtf");
        mimeExtnMap.put("text/html", "html");
        mimeExtnMap.put("message/rfc822", "eml");
        mimeExtnMap.put("image/heic", "heic");
        mimeExtnMap.put("application/zip", "zip");
        mimeExtnMap.put("text/x-vcard", "vcf");
        mimeExtnMap.put("application/pkcs7-signature", "p7s");
        mimeExtnMap.put("video/mp4", "mp4");
        mimeExtnMap.put("application/x-tika-msoffice", "msg");

        return mimeExtnMap;
    }

    public static void copyFile(String sourcePath, String destinationDir) throws IOException {
        File srcfile = new File(sourcePath);
        File destFolder = new File(destinationDir);
        FileUtils.copyFileToDirectory(srcfile, destFolder);
    }

    public void updateResultSheet() throws Exception {
        AmazonS3 s3Client = null;
        Workbook workbook = null;
        try {
            // Instantiate a new Workbook object
            workbook = new Workbook(appConfigs.getExcelPath());

            for (int sheetNo = 7; sheetNo <= 7; sheetNo++) {
//            for (int sheetNo = 0; sheetNo <= 0; sheetNo++) {
                Worksheet worksheet = workbook.getWorksheets().get(sheetNo);
                // Get the cells from the sheet
                Cells cells = worksheet.getCells();

                // Get the maximum data row
                int maxDataRow = cells.getMaxDataRow();


                // Loop through each row
                for (int i = 1; i <= maxDataRow; i++) {
                    // Get the row
                    Row row = cells.getRow(i);

                    // Loop through each cell in the row
                    // Get the cell
                    Cell emailReceivedDateCell = row.get(3);
                    Cell receiptCell = row.get(4);

                    Cell fileNameCell = row.get(5); //editable filename

                    Cell s3PathCell = row.get(6);// s3 path


                    if (receiptCell != null) {
                        if ("".equalsIgnoreCase(s3PathCell.getStringValue())) {
//                        if ("".equalsIgnoreCase(emailReceivedDateCell.getStringValue())) {

                            Row matchRow = getRowFromSheet1(workbook, receiptCell.getStringValue());
                            if (null != matchRow) {
//                                emailReceivedDateCell.setValue(matchRow.get(1).getStringValue());
                                s3PathCell.setValue(matchRow.get(2).getStringValue());
                                log.info("s3path updated.....      ");
                                if ("".equalsIgnoreCase(fileNameCell.getStringValue())) {
                                    String aValue = row.get(8).getStringValue();
                                    if (!aValue.contains(",")) {
                                        fileNameCell.setValue(aValue);
                                        log.info("fileanme updated.....      ");
                                    }
                                }
                            }

                        }

                        if (i % 500 == 0)
                            workbook.save(appConfigs.getExcelPath());
//                            if (i % 2000 == 0)
//                                return;
                    }
                }
            }
        } catch (Exception e) {
            log.error("caught Exception readExcelFileRowByRow", e);
            if (workbook != null) {
                workbook.save(appConfigs.getExcelPath());
                log.info("Excel updated.....      ");
            }
            System.exit(0);
        } finally {
            if (workbook != null) {
                workbook.save(appConfigs.getExcelPath());
                log.info("Excel updated.....      ");
            }
        }
    }

    private Row getRowFromSheet1(Workbook workbook, String receiptNo) {
        Worksheet sheet1 = workbook.getWorksheets().get(2);
        Cells cells = sheet1.getCells();
        int maxDataRow = cells.getMaxDataRow();

        for (int i = 1; i <= maxDataRow; i++) {
            Row row = cells.getRow(i);
            Cell receiptCell = row.getCellOrNull(0);
            if (receiptCell != null && receiptNo.equalsIgnoreCase(receiptCell.getStringValue())) {
                return row;
            }
        }
        return null;
    }

    public void prepareFlattenedFiles() throws Exception {
        AmazonS3 s3Client = null;
        Workbook workbook = null;
        try {
            s3Client = AmazonS3ClientBuilder.standard()
                    .withRegion(appConfigs.getS3RegionName())
                    .withCredentials(DefaultAWSCredentialsProviderChain.getInstance())
                    .build();
            // Instantiate a new Workbook object
            workbook = new Workbook(appConfigs.getExcelPath());

            for (int sheetNo = 1; sheetNo <= 1; sheetNo++) {
//            for (int sheetNo = 0; sheetNo <= 0; sheetNo++) {
                Worksheet worksheet = workbook.getWorksheets().get(sheetNo);
                // Get the cells from the sheet
                Cells cells = worksheet.getCells();

                // Get the maximum data row
                int maxDataRow = cells.getMaxDataRow();


                // Loop through each row
                for (int i = 1; i <= maxDataRow; i++) {
                    // Get the row
                    Row row = cells.getRow(i);

                    // Loop through each cell in the row
                    // Get the cell
                    Cell statusCell = row.get(8);
                    Cell s3PathCell = row.getCellOrNull(11);
                    Cell fileNameCell = row.get(7);
                    Cell rcptNumberCell = row.get(4);

                    if (!"COMPLETED".equalsIgnoreCase(statusCell.getStringValue()) && s3PathCell != null && fileNameCell != null) {
                        String s3Key = s3PathCell.getStringValue();
                        Detail detail = new Detail();
                        try {
                            byte[] emlData;
                            if (appConfigs.isLocal()) {
                                emlData = readFileFromLocal(s3Key, workbook);
                            } else {
                                emlData = readFileFromS3(workbook, s3Configurations.s3Client(), s3Key);
                            }
                            if (null != emlData) {
                                String fileName = null != fileNameCell ? fileNameCell.getStringValue() : "";
                                String receiptNumber = null != rcptNumberCell ? rcptNumberCell.getStringValue() : "";
                                processFile(emlData, detail, fileName, receiptNumber);
                                statusCell.setValue("COMPLETED");
                                log.info("row completed.....      " + i);
                                statusCell.setValue("COMPLETED");
                            } else {
                                log.info("row failed.....      " + i + " File not found in S3");
                                statusCell.setValue("COMPLETED");
                            }
                        } catch (Exception e) {
//                            detail.comments.append(" [" + e.getMessage() + "] ");
//                            writeEditablePdfDetailToExcel(detail, row, "FAILED");
                            log.error("Failed to process row.....      " + i, e.getMessage());
                        }
                    }
                }
            }
        } catch (Exception e) {
            log.error("caught Exception readExcelFileRowByRow", e);
        } finally {
            if (workbook != null) {
                workbook.save(appConfigs.getExcelPath());
                log.info("Excel updated.....      ");
            }
        }
    }

    public void identifyMultilineInWordDoc() throws Exception {
        AmazonS3 s3Client = null;
        Workbook workbook = null;
        try {
            s3Client = AmazonS3ClientBuilder.standard()
                    .withRegion(appConfigs.getS3RegionName())
                    .withCredentials(DefaultAWSCredentialsProviderChain.getInstance())
                    .build();
            // Instantiate a new Workbook object
            workbook = new Workbook(appConfigs.getExcelPath());

            for (int sheetNo = 0; sheetNo <= 0; sheetNo++) {
//            for (int sheetNo = 0; sheetNo <= 0; sheetNo++) {
                Worksheet worksheet = workbook.getWorksheets().get(sheetNo);
                // Get the cells from the sheet
                Cells cells = worksheet.getCells();

                // Get the maximum data row
                int maxDataRow = cells.getMaxDataRow();


                // Loop through each row
                for (int i = 1; i <= maxDataRow; i++) {
                    // Get the row
                    Row row = cells.getRow(i);

                    // Loop through each cell in the row
                    // Get the cell
                    Cell emlPathCell = row.getCellOrNull(11);
                    Cell dbAttachmentsCell = row.get(12);
                    Cell statusCell = row.get(8);

                    if (emlPathCell != null) {
                        String s3Key = emlPathCell.getStringValue();
                        if (!"COMPLETED".equalsIgnoreCase(statusCell.getStringValue()) && StringUtils.isNotBlank(s3Key)) {
                            Detail detail = new Detail();
                            try {
                                byte[] emlData;
                                if (appConfigs.isLocal()) {
                                    emlData = readFileFromLocal(s3Key, workbook);
                                } else {
                                    emlData = readFileFromS3(workbook, s3Configurations.s3Client(), s3Key);
                                }
                                if (null != emlData) {
                                    detail.attachmentsFromDb = null != dbAttachmentsCell ? dbAttachmentsCell.getStringValue() : "";
                                    detail = processPart(emlData, detail, "word");
                                    writeDocDetailToExcel(detail, row, "COMPLETED");
                                } else {
                                    writeDocDetailToExcel(detail, row, "File not found in S3");
                                    log.info("row failed.....      " + i + " File not found in S3");
                                }
                            } catch (Exception e) {
                                detail.comments.append(" [" + e.getMessage() + "] ");
                                writeDocDetailToExcel(detail, row, "FAILED");
                            } finally {
                                writeDocResult(workbook, row, detail);
                            }
                            // Save the workbook
                            log.info("row finished.....      " + i);
                            if (i % 100 == 0)
                                workbook.save(appConfigs.getExcelPath());
//                            if (i % 40000 == 0)
//                                return;
                        }
                    }
                }
            }
        } catch (Exception e) {
            log.error("caught Exception readExcelFileRowByRow", e);
        } finally {
            if (workbook != null) {
                workbook.save(appConfigs.getExcelPath());
                log.info("Excel updated.....      ");
            }
            if (s3Client != null) {
                s3Client.shutdown();
            }
        }
    }


    private void writeDocDetailToExcel(Detail detail, Row row, String status) {
        Cell hasMultiline = row.get(6);
        Cell docFilename = row.get(7);
        Cell statusCell = row.get(8);
        Cell commentsCell = row.get(9);
        Cell allAttachmentNamesCell = row.get(10);

        docFilename.setValue(detail.fileName);
        commentsCell.setValue(detail.comments.toString());
        allAttachmentNamesCell.setValue(detail.allAttachmentsBuilder.toString());
        statusCell.setValue(status);
        hasMultiline.setValue(detail.hasMultiline);

    }

}
