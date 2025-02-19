package com.ontlogieai;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.util.*;

public class ExcelProcessor {

    private static final Logger LOGGER = LoggerFactory.getLogger(ExcelProcessor.class);

    public static void readAndWriteExcelFile(File inputFile, File outputFile) throws IOException {
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

    private static void processFloormanagerSheet(Workbook workbook, Workbook newWorkbook, Workbook refWorkbook) {
        Sheet sheet = workbook.getSheet("Floormanager");
        if (sheet == null) {
            LOGGER.warn("Sheet 'Floormanager' not found.");
            return;
        }

        Sheet newSheet = newWorkbook.getSheet("J270-06-demo");
        if (newSheet == null) {
            newSheet = newWorkbook.createSheet("J270-06-demo");
        }

        addHeaderRow(newSheet);

        for (Row row : sheet) {
            String deviceTag = getDeviceTag(row, 3);
            String pointDescription = getPointDescription(row, 4);
            String standardDeviceTag = getStandardDeviceTagSpecificToDeviceTag(deviceTag, pointDescription);

            LOGGER.debug("Row processed - Device Tag: {}, Point Description: {}, Standard Device Tag: {}",
                    deviceTag, pointDescription, standardDeviceTag);

            if (standardDeviceTag != null) {
                copyRelatedRowsFromRefAlarmAndParameterListSheet(newSheet, refWorkbook, standardDeviceTag, deviceTag);
            }
        }
    }

    private static void copyRelatedRowsFromRefAlarmAndParameterListSheet(Sheet newSheet, Workbook refWorkbook, String standardDeviceTag, String deviceTag) {

        try {
            copyRows(refWorkbook, newSheet, standardDeviceTag, deviceTag);
        } catch (Exception e) {
            LOGGER.error("Error processing reference workbook", e);
        }
    }

    private static void copyRows(Workbook refWorkbook, Sheet newSheet, String standardDeviceTag, String deviceTag) {
        List<Row> matchedRows = getDeviceIdRows(standardDeviceTag,refWorkbook);
        // Get the last row number in newSheet to start appending
        int newRowNum = newSheet.getLastRowNum() + 1;

        for (Row sourceRow : matchedRows) {
            Row newRow = newSheet.createRow(newRowNum++); // Create new row in newSheet

            // Copy each cell from sourceRow to newRow
            for (Cell sourceCell : sourceRow) {
                Cell newCell = newRow.createCell(sourceCell.getColumnIndex());

                // Copy cell value
                switch (sourceCell.getCellType()) {
                    case STRING:
                        // Replace standardDeviceTag with deviceTag in string values
                        String updatedValue = sourceCell.getStringCellValue().replace(standardDeviceTag, deviceTag);
                        newCell.setCellValue(updatedValue);
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
                // Copy cell style
                // newCell.setCellStyle(sourceCell.getCellStyle());
            }
        }
    }

    public static List<Row> getDeviceIdRows(String deviceCode, Workbook refWorkbook) {
        List<Row> matchingRows = new ArrayList<>();

        try {

            Sheet sheet = refWorkbook.getSheetAt(1); // Assuming the data is in the first sheet

            // Get the header row (first row)
            Row headerRow = sheet.getRow(0);

            // Find the column index for "Device Tag"
            int deviceTagColumnIndex = -1;
            for (Cell cell : headerRow) {
                if (cell.getStringCellValue().equalsIgnoreCase("Device Tag")) {
                    deviceTagColumnIndex = cell.getColumnIndex();
                    break;
                }
            }

            // If the column is not found, return an empty list
            if (deviceTagColumnIndex == -1) {
                System.out.println("Device Tag column not found!");
                return matchingRows;
            }

            // Iterate through all rows and check the condition
            Iterator<Row> rowIterator = sheet.iterator();
            rowIterator.next(); // Skip the header row
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Cell deviceTagCell = row.getCell(deviceTagColumnIndex);

                if (deviceTagCell != null && deviceTagCell.getCellType() == CellType.STRING) {
                    String deviceTagValue = deviceTagCell.getStringCellValue();
                    if (null != deviceTagValue && deviceTagValue.contains(deviceCode)) {
                        matchingRows.add(row);
                    }
                }
            }

            refWorkbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

        return matchingRows;
    }

    private static String getStandardDeviceTagSpecificToDeviceTag(String deviceTag, String pointDescription){


        String keyPrifix = getDeviceType(deviceTag);
        String deviceKeyPostfix = getDeviceKey(pointDescription);
        String deviceKey = keyPrifix +"-"+ deviceKeyPostfix;

        Map<String, String> standardDeviceTag = new HashMap<>();
        //TT
        standardDeviceTag.put("TT-Potable Water - Temperature", "J100-06-2TT-001");
        standardDeviceTag.put("TT-Potable  Hot Water", "J100-06-2TT-002");
        standardDeviceTag.put("TT-Non Potable Water","J130-06-2TT-001");
        standardDeviceTag.put("TT-Chilled Water - Supply Temperature", "J460-01-2TT-612");
        standardDeviceTag.put("TT-Chilled Water - Return Temperature", "J460-01-2TT-613");
        standardDeviceTag.put("TT-Supply Air", "J460-01-2TT-614");
        standardDeviceTag.put("TT-Return Air", "J460-01-2TT-616");

        //FT
        standardDeviceTag.put("FT-Chilled Water", "J460-01-2FT-601");
        standardDeviceTag.put("FT-Hot Water", "J460-01-2FT-602");
        standardDeviceTag.put("FT-Supply Air Flow", "J460-01-2FT-603");
        standardDeviceTag.put("FT-Return Air Flow", "J460-01-2FT-604");
        standardDeviceTag.put("FT-Potable Water", "J100-06-2FT-001");
        standardDeviceTag.put("FT- Non Potable Water", "J130-06-2FT-001");
        standardDeviceTag.put("FT-Compressed Air", "J305-06-2FT-001");
        standardDeviceTag.put("FT-Carbon Dioxide Gas","J305-06-2FT-001");
        standardDeviceTag.put("FT-Nitrogen Gas","J330-06-2FT-001");
        standardDeviceTag.put("FT-Demi Water","J140-06-2FT-001");

        return standardDeviceTag.get(deviceKey);
    }


    private static String getDeviceKey(String deviceTag){
        if(deviceTag == null){
            return "";
        }

        if(deviceTag.contains("Potable Water - Temperature")){
            return "Potable Water - Temperature";
        }
        if (deviceTag.contains("Potable  Hot Water")){
            return "Potable  Hot Water";
        }
        if (deviceTag.contains("Non Potable Water")){
            return "Non Potable Water";
        }
        if (deviceTag.contains("Chilled Water - Supply Temperature")){
            return "Chilled Water - Supply Temperature";
        }
        if (deviceTag.contains("Chilled Water - Return Temperature")){
            return "Chilled Water - Return Temperature";
        }
        if (deviceTag.contains("Supply Air")){
            return "Supply Air";
        }
        if (deviceTag.contains("Return Air")){
            return "Return Air";
        }

        if (deviceTag.contains("Chilled Water")) {
            return "Chilled Water";
        }
        if (deviceTag.contains("Hot Water")) {
            return "Hot Water";
        }
        if (deviceTag.contains("Supply Air Flow")) {
            return "Supply Air Flow";
        }
        if (deviceTag.contains("Return Air Flow")) {
            return "Return Air Flow";
        }
        if (deviceTag.contains("Potable Water")) {
            return "Potable Water";
        }

        if (deviceTag.contains("Compressed Air")) {
            return "Compressed Air";
        }
        if (deviceTag.contains("Carbon Dioxide Gas")) {
            return "Carbon Dioxide Gas";
        }
        if (deviceTag.contains("Nitrogen Gas")) {
            return "Nitrogen Gas";
        }
        if (deviceTag.contains("Demi Water")) {
            return "Demi Water";
        }
        return "";
    }

    private static String getDeviceType(String deviceTag){
        if(deviceTag == null){
            return "";
        }

        if(deviceTag.contains("TT")){
            return "TT";
        }
        if (deviceTag.contains("FT")){
            return "FT";
        }
        return "";
    }

    private static String getDeviceTag(Row row, int index) {
        Cell cell = row.getCell(index);
        if (cell != null) {
            return switch (cell.getCellType()) {
                case STRING -> cell.getStringCellValue();
                case NUMERIC -> String.valueOf(cell.getNumericCellValue());
                case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
                case FORMULA -> cell.getCellFormula();
                default -> "";
            };
        }
        return "";
    }

    private static String getPointDescription(Row row, int index) {
        Cell cell = row.getCell(index);
        if (cell != null) {
            return switch (cell.getCellType()) {
                case STRING -> cell.getStringCellValue();
                case NUMERIC -> String.valueOf(cell.getNumericCellValue());
                case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
                case FORMULA -> cell.getCellFormula();
                default -> "";
            };
        }
        return "";
    }

    private static void addHeaderRow(Sheet sheet) {
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
    }

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
            Row newRow = targetSheet.createRow(sourceRow.getRowNum());
            copyRow(sourceRow, newRow);
        }
    }

    public static void copyRow(Row sourceRow, Row targetRow) {

        // Iterate over the cells in the source row
        for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
            // Get the cell from the source row
            Cell sourceCell = sourceRow.getCell(i);

            // If the source cell is null, continue
            if (sourceCell == null) continue;

            // Create a cell in the target row
            Cell targetCell = targetRow.createCell(i);

            // Copy the value and style from the source cell to the target cell
            copyCell(sourceCell, targetCell);
        }
    }

    private static void copyCell(Cell sourceCell, Cell targetCell) {
        // Copy the cell value
        switch (sourceCell.getCellType()) {
            case STRING:
                targetCell.setCellValue(sourceCell.getStringCellValue());
                break;
            case NUMERIC:
                targetCell.setCellValue(sourceCell.getNumericCellValue());
                break;
            case BOOLEAN:
                targetCell.setCellValue(sourceCell.getBooleanCellValue());
                break;
            case FORMULA:
                targetCell.setCellFormula(sourceCell.getCellFormula());
                break;
            case BLANK:
                targetCell.setBlank();
                break;
            default:
                break;
        }
    }
}
