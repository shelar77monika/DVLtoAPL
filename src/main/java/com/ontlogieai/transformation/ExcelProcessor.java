package com.ontlogieai.transformation;

import com.ontlogieai.Main;
import com.ontlogieai.config.Configuration;
import com.ontlogieai.config.ConfigurationReader;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.util.*;

public class ExcelProcessor {

    private static final Logger LOGGER = LoggerFactory.getLogger(ExcelProcessor.class);

    private static final String[] HEADERS = {
            "Rev Nr", "Nr", "Outstation", "Device Tag", "Function", "Point Description", "EBI Tag", "JACE Tag",
            "Range (Low) / State 0", "Range (High) / State 1", "State 2", "State 3", "State 4", "State 5", "State 6",
            "State 7", "State 8", "State 9", "State 16", "State 32", "State 64", "State 128", "State 8192",
            "State 16384", "State 32768", "Delay Timer (Sec)", "Hysteresis", "Control Level",
            "Electronic Signature Type", "Unit", "Setting", "Controller Alarm Tag", "Alarm Type",
            "Reset", "Remarks"
    };

    private final DeviceTagMapper deviceTagMapper;
    private final Configuration configuration;

    private int deviceTagColumnIndex = -1;
    private int pointDescriptorColumnIndex = -1;

    public ExcelProcessor(){
        deviceTagMapper = new DeviceTagMapper();
        configuration = ConfigurationReader.getConfiguration();
    }

    public void readAndWriteExcelFile(File inputFile, File outputFile) throws IOException {
        LOGGER.info("Reading Excel file: {}", inputFile.getName());

        try (Workbook workbook = WorkbookFactory.createWorkbook(inputFile);
             Workbook newWorkbook = new XSSFWorkbook();
             FileOutputStream fos = new FileOutputStream(outputFile);
             InputStream refFis = getReferenceFileStream();
             Workbook refWorkbook = new XSSFWorkbook(refFis)) {

            copyReferenceSheets(refWorkbook, newWorkbook);
            processFloormanagerSheet(workbook, newWorkbook, refWorkbook);
            newWorkbook.write(fos);

        } catch (Exception e) {
            LOGGER.error("Error processing Excel file", e);
            throw e;
        }
    }

    static class WorkbookFactory {
        static Workbook createWorkbook(File file) throws IOException {
            try (FileInputStream fis = new FileInputStream(file)) {
                return file.getName().endsWith(".xlsx") ? new XSSFWorkbook(fis) : new HSSFWorkbook(fis);
            }
        }
    }

    private static Sheet getOrCreateSheet(Workbook workbook, String sheetName) {
        Sheet sheet = workbook.getSheet(sheetName);
        return (sheet != null) ? sheet : workbook.createSheet(sheetName);
    }

    private  void processFloormanagerSheet(Workbook workbook, Workbook newWorkbook, Workbook refWorkbook) {
        Sheet sheet = workbook.getSheet("Floormanager");
        if (sheet == null) {
            LOGGER.warn("Sheet 'Floormanager' not found.");
            return;
        }

        Sheet newSheet = getOrCreateSheet(newWorkbook, "J270-06-demo");
        addHeaderRow(newSheet);

        deviceTagColumnIndex = getColumnIndex(sheet, "Device Tag");
        pointDescriptorColumnIndex = getColumnIndex(sheet, "Point Descriptor");

        if(deviceTagColumnIndex != -1 && pointDescriptorColumnIndex != -1){
            for (Row row : sheet) {
                processRow(row, newSheet, refWorkbook, deviceTagColumnIndex, pointDescriptorColumnIndex);
            }
        }else{
            LOGGER.info("Excel can not be processed cause Device Tag and Point Descriptor columns not found in the excel");
        }

    }

    private  void processRow(Row row, Sheet newSheet, Workbook refWorkbook, int deviceTagColumnIndex, int pointDescriptorColumnIndex) {
        if (isRowEmpty(row)) {
            LOGGER.debug("Skipping empty row");
            return;
        }

        String deviceTag = getCellValue(row, deviceTagColumnIndex);
        String pointDescription = getCellValue(row, pointDescriptorColumnIndex);
        String standardDeviceTag = deviceTagMapper.getStandardDeviceTag(deviceTag, pointDescription);

        LOGGER.debug("Processing row - Device Tag: {}, Point Description: {}, Standard Device Tag: {}",
                deviceTag, pointDescription, standardDeviceTag);

        if (standardDeviceTag != null && !standardDeviceTag.equalsIgnoreCase("Device Tag") && !standardDeviceTag.isEmpty()) {
            copyRowsFromRefSheet(newSheet, refWorkbook, standardDeviceTag, deviceTag, pointDescription);
        }
    }

    private boolean isRowEmpty(Row row) {
        if (row == null) return true;
        for (Cell cell : row) {
            if (cell.getCellType() != CellType.BLANK && !cell.toString().trim().isEmpty()) {
                return false;
            }
        }
        return true;
    }

    private void copyRowsFromRefSheet(Sheet newSheet, Workbook refWorkbook, String standardDeviceTag, String deviceTag, String pointDescription) {
        try {
            List<Row> matchedRows = getMatchingRows(refWorkbook, standardDeviceTag);
            copyRows(newSheet, matchedRows, standardDeviceTag, deviceTag, pointDescription);
        } catch (Exception e) {
            LOGGER.error("Error processing reference workbook", e);
        }
    }

    private  List<Row> getMatchingRows(Workbook refWorkbook, String standardDeviceTag) {
        List<Row> matchingRows = new ArrayList<>();
        Sheet sheet = refWorkbook.getSheetAt(1);
        //int deviceTagColumnIndex = getColumnIndex(sheet, "Device Tag");
        if (deviceTagColumnIndex == -1) return matchingRows;

        for (Row row : sheet) {
            String cellValue = getCellValue(row, deviceTagColumnIndex);
            if (cellValue.equalsIgnoreCase(standardDeviceTag)) {
                matchingRows.add(row);
            }
        }
        return matchingRows;
    }

    private void copyRows(Sheet newSheet, List<Row> sourceRows, String standardDeviceTag, String deviceTag, String pointDescription) {
        int newRowNum = newSheet.getLastRowNum() + 1;
        for (Row sourceRow : sourceRows) {
            Row newRow = newSheet.createRow(newRowNum++);
            copyRowData(sourceRow, newRow, standardDeviceTag, deviceTag, pointDescription);
        }
    }

    private  void copyRowData(Row sourceRow, Row newRow, String standardDeviceTag, String deviceTag, String pointDescription) {
        for (Cell sourceCell : sourceRow) {
            Cell newCell = newRow.createCell(sourceCell.getColumnIndex());

            switch (sourceCell.getCellType()) {
                case STRING:
                    if(sourceCell.getColumnIndex() == configuration.getDeviceTagIndex()){
                        newCell.setCellValue(sourceCell.getStringCellValue().replace(standardDeviceTag, deviceTag));
                    }else if(sourceCell.getColumnIndex() == configuration.getPointDescriptionIndexInOutputFile()){
                        String regexPointDescriptor = getRegexPointDescriptor(pointDescription, deviceTag);
                        String replacedString = replaceFirstPart(regexPointDescriptor,pointDescription, sourceCell.getStringCellValue());
                        newCell.setCellValue(replacedString);
                    }else {
                        newCell.setCellValue(sourceCell.getStringCellValue().replace(standardDeviceTag, deviceTag));
                    }
                    break;
                case NUMERIC:
                    newCell.setCellValue(sourceCell.getNumericCellValue());
                    break;
                case BOOLEAN:
                    newCell.setCellValue(sourceCell.getBooleanCellValue());
                    break;
                case FORMULA:
                    newCell.setCellFormula(sourceCell.getCellFormula());
                    break;
                case BLANK:
                    newCell.setCellType(CellType.BLANK);
                    break;
                default:
                    break;
            }
        }
    }

    public static String replaceFirstPart(String pointDescriptorRegexValue, String source, String target) {
        // Build a regex pattern using the dynamic pointDescriptorRegexValue
        String regex = "^(.*? - " + pointDescriptorRegexValue + ")";

        // Extracting the matching prefix from source
        String newPrefix = source.replaceAll(regex, "$1");

        // Replacing the first part in target string
        return target.replaceFirst(regex, newPrefix);
    }


    private String getRegexPointDescriptor(String pointDescription, String deviceTag) {
        String regex = "";
        String deviceType = DeviceTagUtility.getDeviceType(deviceTag);
        List<String> pointDescriptorsList = configuration.getDeviceKeyMapping().get(deviceType);
        if(null != pointDescriptorsList && !pointDescriptorsList.isEmpty()){
            regex = pointDescriptorsList.stream()
                    .filter(pointDescription::contains)
                    .findFirst()
                    .orElse("");
        }
        return regex;
    }

    private String getCellValue(Row row, int columnIndex) {
        Cell cell = row.getCell(columnIndex);
        return (cell != null) ? cell.getStringCellValue() : "";
    }

    private int getColumnIndex(Sheet sheet, String columnName) {
        Row headerRow = getHeaderRow(sheet, configuration.getRequiredHeaders());
        if(null != headerRow){
            for (Cell cell : headerRow) {
                if (cell.getStringCellValue().equalsIgnoreCase(columnName)) {
                    return cell.getColumnIndex();
                }
            }
        }
        return -1;
    }

    private Row getHeaderRow(Sheet sheet, List<String> requiredHeaders) {
        Row headerRow = sheet.getRow(0);

        if (isRowEmpty(headerRow) || containsRequiredHeaders(headerRow, requiredHeaders)) {
            LOGGER.debug("Skipping header row at index {} cause it's empty or missing required headers", 0);
            LOGGER.debug("Trying to find headers at row index {}", 1);

            headerRow = sheet.getRow(1);
            if (isRowEmpty(headerRow) || containsRequiredHeaders(headerRow, requiredHeaders)) {
                LOGGER.error("Did not find the required headers in the Excel file");
                return null;
            }
        }
        return headerRow;
    }

    private boolean containsRequiredHeaders(Row row, List<String> requiredHeaders) {
        Set<String> headersInRow = new HashSet<>();
        for (Cell cell : row) {
            headersInRow.add(cell.toString().trim());
        }
        return !headersInRow.containsAll(requiredHeaders);
    }

    public static void addHeaderRow(Sheet sheet) {
        Row headerRow = sheet.createRow(0);
        CellStyle headerStyle = createHeaderStyle(sheet.getWorkbook());

        for (int i = 0; i < HEADERS.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(HEADERS[i]);
            cell.setCellStyle(headerStyle);
            sheet.autoSizeColumn(i);
        }
    }

    private static CellStyle createHeaderStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        return style;
    }


    /*private static void addHeaderRow(Sheet sheet) {
        Row headerRow = sheet.createRow(0); // Create the first row (header row)

        String[] headers = {"Rev Nr","Nr","Outstation",	"Device Tag","Function","Point Description","EBI Tag","JACE Tag","Range (Low) / State 0","Range (High) / State 1","State 2","State 3","State 4","State 5","State 6","State 7","State 8","State 9","State 16","State 32","State 64","State 128","State 8192","State 16384","State 32768","Delay Timer (Sec)","Hysteresis","Control Level","Electronic Signature Type","Unit","Setting","Controller Alarm Tag","	Alarm Type","Reset","Remarks"
        };

        // Create header style
        CellStyle headerStyle = sheet.getWorkbook().createCellStyle();
        org. apache. poi. ss. usermodel.Font headerFont = sheet.getWorkbook().createFont();
        headerFont.setBold(true);
        headerStyle.setFont(headerFont);

        // Add headers to the first row
        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(headerStyle);
            sheet.autoSizeColumn(i); // Adjust column width automatically
        }
    }*/

    private static InputStream getReferenceFileStream() throws IOException {
        InputStream refFis = Main.class.getClassLoader().getResourceAsStream("J270-06-Alarm-And-Parameter-list.xlsx");
        if (refFis == null) {
            throw new IOException("Reference Excel file not found in resources.");
        }
        return refFis;
    }
    private static void copyReferenceSheets(Workbook refWorkbook, Workbook newWorkbook) {
        for (Sheet refSheet : refWorkbook) {
            if (!"J270-06".equalsIgnoreCase(refSheet.getSheetName())) {
                Sheet newSheet = newWorkbook.createSheet(refSheet.getSheetName());
                copySheet(refSheet, newSheet);
            }
        }
    }

    private static void copySheet(Sheet sourceSheet, Sheet targetSheet) {
        for (Row sourceRow : sourceSheet) {
            Row targetRow = targetSheet.createRow(sourceRow.getRowNum());
            copyRow(sourceRow, targetRow);
        }
    }

    private static void copyRow(Row sourceRow, Row targetRow) {
        for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
            Cell sourceCell = sourceRow.getCell(i);
            if (sourceCell != null) {
                Cell targetCell = targetRow.createCell(i);
                copyCell(sourceCell, targetCell);
            }
        }
    }

    private static void copyCell(Cell sourceCell, Cell targetCell) {
        switch (sourceCell.getCellType()) {
            case STRING -> targetCell.setCellValue(sourceCell.getStringCellValue());
            case NUMERIC -> targetCell.setCellValue(sourceCell.getNumericCellValue());
            case BOOLEAN -> targetCell.setCellValue(sourceCell.getBooleanCellValue());
            case FORMULA -> targetCell.setCellFormula(sourceCell.getCellFormula());
            case BLANK -> targetCell.setBlank();
            default -> {
            }
        }
    }
}
