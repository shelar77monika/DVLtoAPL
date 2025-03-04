package com.ontlogieai;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRow;
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

        int deviceTagColumnIndex = getDeviceTagColumnIndex(sheet);
        int pointDescriptorColumnIndex = getPointDescriptorColumnIndex(sheet);

        for (Row row : sheet) {
            String deviceTag = getDeviceTag(row, deviceTagColumnIndex);
            String pointDescription = getPointDescription(row, pointDescriptorColumnIndex);
            String standardDeviceTag = getStandardDeviceTagSpecificToDeviceTag(deviceTag, pointDescription);

            LOGGER.debug("Row processed - Device Tag: {}, Point Description: {}, Standard Device Tag: {}",
                    deviceTag, pointDescription, standardDeviceTag);

            if(deviceTag.equalsIgnoreCase("J460-02-2TT-717")){
                if (standardDeviceTag != null) {
                    copyRelatedRowsFromRefAlarmAndParameterListSheet(newSheet, refWorkbook, standardDeviceTag, deviceTag);
                }
            }else if (standardDeviceTag != null) {
                copyRelatedRowsFromRefAlarmAndParameterListSheet(newSheet, refWorkbook, standardDeviceTag, deviceTag);
            }
        }
    }

    private static int getDeviceTagColumnIndex(Sheet sheet){
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
        return deviceTagColumnIndex;
    }

    private static int getPointDescriptorColumnIndex(Sheet sheet){
        // Get the header row (first row)
        Row headerRow = sheet.getRow(0);

        // Find the column index for "Device Tag"
        int pointDescriptorColumnIndex = -1;
        for (Cell cell : headerRow) {
            if (cell.getStringCellValue().equalsIgnoreCase("Point Descriptor")) {
                pointDescriptorColumnIndex = cell.getColumnIndex();
                break;
            }
        }
        return pointDescriptorColumnIndex;
    }

    private static void copyRelatedRowsFromRefAlarmAndParameterListSheet(Sheet newSheet, Workbook refWorkbook, String standardDeviceTag, String deviceTag) {

        try {
            copyRows(refWorkbook, newSheet, standardDeviceTag, deviceTag);
        } catch (Exception e) {
            LOGGER.error("Error processing reference workbook", e);
        }
    }

    private static void copyRows(Workbook refWorkbook, Sheet newSheet, String standardDeviceTag, String deviceTag) {
        List<Row> matchedRows = getDeviceIdRows(standardDeviceTag,refWorkbook, deviceTag);
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

    public static List<Row> getDeviceIdRows(String standardDeviceTag, Workbook refWorkbook, String deviceCode) {
        List<Row> matchingRows = new ArrayList<>();

        try {

            Sheet sheet = refWorkbook.getSheetAt(1); // Assuming the data is in the first sheet

            // Find the column index for "Device Tag"
            int deviceTagColumnIndex = getDeviceTagColumnIndex(sheet);

            // If the column is not found, return an empty list
            if (deviceTagColumnIndex == -1) {
                System.out.println("Device Tag column not found!");
                return matchingRows;
            }

            // Iterate through all rows and check the condition
            Iterator<Row> rowIterator = sheet.iterator();
            LOGGER.info("Row Count: {}", sheet.getLastRowNum());
            rowIterator.next(); // Skip the header row
            int counter = 1;
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Cell deviceTagCell = row.getCell(deviceTagColumnIndex);

                if (deviceTagCell != null && deviceTagCell.getCellType() == CellType.STRING) {
                    String deviceTagValue = deviceTagCell.getStringCellValue();
                    LOGGER.info("deviceTagValue : {} standardDeviceTag {} count {}", deviceTagValue, standardDeviceTag, counter );
                    if (null != deviceTagValue && deviceTagValue.equalsIgnoreCase(standardDeviceTag)) {
                        // Replace the value with deviceCode
                        //deviceTagCell.setCellValue(deviceCode);
                        matchingRows.add(row);
                    }
                    counter++;
                }
            }

//            for (Row row: matchingRows){
//                Cell deviceTagCell = row.getCell(deviceTagColumnIndex);
//                if (null != deviceTagCell) {
//                    deviceTagCell.setCellValue(deviceCode);
//                }
//            }
            refWorkbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }


        return matchingRows;
    }



    private static String getStandardDeviceTagSpecificToDeviceTag(String deviceTag, String pointDescription){


        if(deviceTag.equalsIgnoreCase("J460-02-2TT-717")){
            LOGGER.info("Here we need to debug...");

        }
        String keyPrefix = getDeviceType(deviceTag);
        String deviceKeyPostfix = getDeviceKey(keyPrefix, pointDescription);
        String deviceKey = keyPrefix +"-"+ deviceKeyPostfix;

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
        standardDeviceTag.put("FT-Non Potable Water", "J130-06-2FT-001");
        standardDeviceTag.put("FT-Compressed Air", "J305-06-2FT-001");
        standardDeviceTag.put("FT-Carbon Dioxide Gas","J305-06-2FT-001");
        standardDeviceTag.put("FT-Nitrogen Gas","J330-06-2FT-001");
        standardDeviceTag.put("FT-Demi Water","J140-06-2FT-001");

        //MT
        standardDeviceTag.put("MT-Supply Air Humidity", "J460-01-2MT-614");
        standardDeviceTag.put("MT-Compressed Air", "J305-06-2MT-001");
        standardDeviceTag.put("MT-Compressed Air - Dewpoint", "J305-06-2MT-001");
        standardDeviceTag.put("MT-Humidity", "J460-01-2MT-601");
        standardDeviceTag.put("MT-Chilled Water Valve - Controller", "J460-01-2FCV-643");

        //PT
        standardDeviceTag.put("PT-Non Potable Water - Pressure", "J130-06-2PT-001");
        standardDeviceTag.put("PT-Compressed Air - Pressure", "J305-06-2PT-001");
        standardDeviceTag.put("PT-Carbon Dioxide Gas - Pressure", "J320-06-2PT-001");
        standardDeviceTag.put("PT-Nitrogen Gas - Pressure", "J330-06-2PT-001");
        standardDeviceTag.put("PT-Demi Water - Inlet Pressure", "J140-06-2PT-001");
        standardDeviceTag.put("PT-Demi Water - Outlet Pressure", "J140-06-2PT-002");
        standardDeviceTag.put("PT-Demi Water - Return Pressure", "J140-06-2PT-003");
        standardDeviceTag.put("PT-Pressure", "J460-02-2PT-902");

        //ACU
        standardDeviceTag.put("ACU-Fan Speed", "J460-02-1ACU-601");
        standardDeviceTag.put("ACU-Fan Coil Unit Control", "J460-01-1ACU-618");

        //XC
        standardDeviceTag.put("XC-Exhaust Fan", "J460-02-2B-902");

        //XT
        standardDeviceTag.put("XT-Occupied", "J460-01-2XT-002");
        standardDeviceTag.put("XT-CO2 Concentration", "J460-01-2XT-001");

        //XA
        standardDeviceTag.put("XA-Thermal Fault Signal", "J270-06-2XA-001");
        standardDeviceTag.put("XA-Surge Voltage Arrester Signal", "J270-06-2XA-004");
        standardDeviceTag.put("XA-Common Fire Alarm", "J270-06-2XA-005");
        standardDeviceTag.put("XA-Circuit Breaker Tripped", "J229-06-2XA-101");
        standardDeviceTag.put("XA-Voltage Surge Arrestor", "J229-06-2XA-102");
        standardDeviceTag.put("XA-UPS Alarm", "J460-01-2XA-802");

        //FCV
        standardDeviceTag.put("FCV-Reheater Valve Control", "J460-02-2FCV-623");
        standardDeviceTag.put("FCV-Heating Valve Control", "J460-01-2FCV-002");
        standardDeviceTag.put("FCV-Cooling Valve Control", "J460-01-2FCV-006");
        standardDeviceTag.put("FCV-Chilled Water Valve", "J460-01-2FCV-643");

        //KS
        standardDeviceTag.put("KS-Labs Day Extension Timer - Timer","J460-02-2KS-603");

        //XI
        standardDeviceTag.put("XI-Labs Day Extension Timer - Indicator","J460-02-2XI-603");

        //PMP
        standardDeviceTag.put("PMP-Chilled Water Circulation Pump","J460-01-1PMP-601");

        //QIT
        standardDeviceTag.put("QIT-Energy Meter", "J229-06-1QIT-001");

        //UPS
        standardDeviceTag.put("UPS-UPS", "J232-06-1UPS-001");

        //VAV
        standardDeviceTag.put("VAV-Return Air Flow Control", "J460-01-1VAV-001");
        standardDeviceTag.put("VAV-Supply Air Flow Control", "J460-01-1VAV-603");
        standardDeviceTag.put("VAV-Fume hood", "J460-02-1VAV-606");
        standardDeviceTag.put("VAV-Air Fow Control", "J460-02-1VAV-607");

        //TC
        standardDeviceTag.put("TC-Room Controller","J460-01-2TC-601");

        //XCV
        standardDeviceTag.put("XCV-Legionella Dump Valve", "J100-06-2XCV-001");

        return standardDeviceTag.get(deviceKey);
    }


    private static String getDeviceKey(String keyPrefix, String pointDescriptor){
        if(pointDescriptor == null){
            return "";
        }

        if(keyPrefix.equalsIgnoreCase("TT")){
            //TT
            if(pointDescriptor.contains("Potable Water - Temperature")){
                return "Potable Water - Temperature";
            }
            if (pointDescriptor.contains("Potable  Hot Water")){
                return "Potable  Hot Water";
            }
            if (pointDescriptor.contains("Non Potable Water")){
                return "Non Potable Water";
            }
            if (pointDescriptor.contains("Chilled Water - Supply Temperature")){
                return "Chilled Water - Supply Temperature";
            }
            if (pointDescriptor.contains("Chilled Water - Return Temperature")){
                return "Chilled Water - Return Temperature";
            }
            if (pointDescriptor.contains("Supply Air")){
                return "Supply Air";
            }
            if (pointDescriptor.contains("Return Air")){
                return "Return Air";
            }
        }

        if(keyPrefix.equalsIgnoreCase("FT")){
            //FT
            if (pointDescriptor.contains("Chilled Water")) {
                return "Chilled Water";
            }
            if (pointDescriptor.contains("Hot Water")) {
                return "Hot Water";
            }
            if (pointDescriptor.contains("Supply Air Flow")) {
                return "Supply Air Flow";
            }
            if (pointDescriptor.contains("Return Air Flow")) {
                return "Return Air Flow";
            }
            if (pointDescriptor.contains("Potable Water")) {
                return "Potable Water";
            }
            if (pointDescriptor.contains("Compressed Air")) {
                return "Compressed Air";
            }
            if (pointDescriptor.contains("Carbon Dioxide Gas")) {
                return "Carbon Dioxide Gas";
            }
            if (pointDescriptor.contains("Nitrogen Gas")) {
                return "Nitrogen Gas";
            }
            if (pointDescriptor.contains("Demi Water")) {
                return "Demi Water";
            }
        }


        if(keyPrefix.equalsIgnoreCase("MT")){
            // MT - Miscellaneous Measurement Tags
            if (pointDescriptor.contains("Supply Air Humidity")) {
                return "Supply Air Humidity";
            }
            if (pointDescriptor.contains("Humidity")) {
                return "Humidity";
            }
            if (pointDescriptor.contains("Chilled Water Valve - Controller")) {
                return "Chilled Water Valve - Controller";
            }
            if(pointDescriptor.contains("Compressed Air - Dewpoint")){
                return "Compressed Air - Dewpoint";
            }
        }

        if(keyPrefix.equalsIgnoreCase("PT")){
            // PT - Pressure Tags
            if(pointDescriptor.contains("Non Potable Water - Pressure")){
                return "Non Potable Water - Pressure";
            }
            if(pointDescriptor.contains("Compressed Air - Pressure")){
                return "Compressed Air - Pressure";
            }
            if(pointDescriptor.contains("Carbon Dioxide Gas - Pressure")){
                return "Carbon Dioxide Gas - Pressure";
            }
            if(pointDescriptor.contains("Nitrogen Gas - Pressure")){
                return "Nitrogen Gas - Pressure";
            }
            if(pointDescriptor.contains("Demi Water - Inlet Pressure")){
                return "Demi Water - Inlet Pressure";
            }
            if(pointDescriptor.contains("Demi Water - Outlet Pressure")){
                return "Demi Water - Outlet Pressure";
            }
            if(pointDescriptor.contains("Demi Water - Return Pressure")){
                return "Demi Water - Return Pressure";
            }
            if(pointDescriptor.contains("Pressure")){
                return "Pressure";
            }
        }


        if(keyPrefix.equalsIgnoreCase("ACU")){
            // ACU - Air Conditioning Unit
            if (pointDescriptor.contains("Fan Coil Unit Control")) {
                return "Fan Coil Unit Control";
            }
            if(pointDescriptor.contains("Fan Speed")){
                return "Fan Speed";
            }
        }

        if(keyPrefix.equalsIgnoreCase("XC")){
            // XC - Exhaust Fan
            if (pointDescriptor.contains("Exhaust Fan")) {
                return "Exhaust Fan";
            }
        }

        if(keyPrefix.equalsIgnoreCase("XT")){
            //XT
            if(pointDescriptor.contains("Occupied")){
                return "Occupied";
            }
            if (pointDescriptor.contains("CO2 Concentration")){
                return "CO2 Concentration";
            }
        }

        if (keyPrefix.equalsIgnoreCase("XA")) {
            // XT - Extended Transmitter (XA category)
            if (pointDescriptor.contains("Thermal Fault Signal")) {
                return "Thermal Fault Signal";
            }
            if (pointDescriptor.contains("Surge Voltage Arrester Signal")) {
                return "Surge Voltage Arrester Signal";
            }
            if (pointDescriptor.contains("Common Fire Alarm")) {
                return "Common Fire Alarm";
            }
            if (pointDescriptor.contains("Circuit Breaker Tripped")) {
                return "Circuit Breaker Tripped";
            }
            if (pointDescriptor.contains("Voltage Surge Arrestor")) {
                return "Voltage Surge Arrestor";
            }
            if (pointDescriptor.contains("UPS Alarm")) {
                return "UPS Alarm";
            }
        }

        if(keyPrefix.equalsIgnoreCase("FCV")){
            if(pointDescriptor.contains("Reheater Valve Control")){
                return "Reheater Valve Control";
            }
            if(pointDescriptor.contains("Heating Valve Control")){
                return "Heating Valve Control";
            }
            if(pointDescriptor.contains("Cooling Valve Control")){
                return "Cooling Valve Control";
            }
            if(pointDescriptor.contains("Chilled Water Valve")){
                return "Chilled Water Valve";
            }
        }
        if(keyPrefix.equalsIgnoreCase("KS")){
            if (pointDescriptor.contains("Labs Day Extension Timer - Timer")){
                return "Labs Day Extension Timer - Timer";
            }
        }


        if(keyPrefix.equalsIgnoreCase("XI")){
            if (pointDescriptor.contains("Labs Day Extension Timer - Indicator")){
                return "Labs Day Extension Timer - Indicator";
            }
        }

        if(keyPrefix.equalsIgnoreCase("PMP")){
            if (pointDescriptor.contains("Chilled Water Circulation Pump")){
                return "Chilled Water Circulation Pump";
            }
        }

        if(keyPrefix.equalsIgnoreCase("QIT")){
            if (pointDescriptor.contains("Energy Meter")){
                return "Energy Meter";
            }
        }
        if(keyPrefix.equalsIgnoreCase("UPS")){
            if (pointDescriptor.contains("UPS")){
                return "UPS";
            }
        }

        if(keyPrefix.equalsIgnoreCase("VAV")){
            if (pointDescriptor.contains("Return Air Flow Control")){
                return "Return Air Flow Control";
            }
            if (pointDescriptor.contains("Supply Air Flow Control")){
                return "Supply Air Flow Control";
            }
            if (pointDescriptor.contains("Fume hood")){
                return "Fume hood";
            }
            if (pointDescriptor.contains("Air Fow Control")){
                return "Air Fow Control";
            }
        }

        if(keyPrefix.equalsIgnoreCase("TC")){
            if (pointDescriptor.contains("Room Controller")){
                return "Room Controller";
            }
        }

        if(keyPrefix.equalsIgnoreCase("XCV")){
            if (pointDescriptor.contains("Legionella Dump Valve")){
                return "Legionella Dump Valve";
            }
        }

        return "";
    }

    private static String getDeviceType(String deviceTag) {
        if (deviceTag == null) {
            return "";
        }

        if (deviceTag.contains("TT")) {
            return "TT"; // Temperature Transmitter
        }
        if (deviceTag.contains("FT")) {
            return "FT"; // Flow Transmitter
        }
        if (deviceTag.contains("MT")) {
            return "MT"; // Miscellaneous Measurement Transmitter
        }
        if (deviceTag.contains("PT")) {
            return "PT"; // Pressure Transmitter
        }
        if (deviceTag.contains("ACU")) {
            return "ACU"; // Air Conditioning Unit
        }
        if (deviceTag.contains("XC")) {
            return "XC"; // Exhaust Controller
        }
        if (deviceTag.contains("XT")) {
            return "XT"; // Exhaust Controller
        }

        if (deviceTag.contains("XA")) {
            return "XA"; // Thermal Fault
        }
        if (deviceTag.contains("FCV")) {
            return "FCV"; // Thermal Fault
        }
        if(deviceTag.contains("KS")){
            return "KS";
        }
        if(deviceTag.contains("PMP")){
            return "PMP";
        }
        if(deviceTag.contains("QIT")){
            return "QIT";
        }
        if(deviceTag.contains("UPS")){
            return "UPS";
        }
        if(deviceTag.contains("VAV")){
            return "VAV";
        }
        if(deviceTag.contains("TC")){
            return "TC";
        }
        if(deviceTag.contains("XCV")){
            return "XCV";
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
