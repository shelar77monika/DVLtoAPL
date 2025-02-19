package com.ontlogieai;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.swing.*;
import java.io.File;
import java.io.IOException;

public class FileProcessor {

    private static final Logger LOGGER = LoggerFactory.getLogger(FileProcessor.class);

    public static void handleFileUpload(JFrame frame) {
        JFileChooser fileChooser = new JFileChooser();
        if (fileChooser.showOpenDialog(frame) == JFileChooser.APPROVE_OPTION) {
            processFile(fileChooser.getSelectedFile());
        }
    }

    public static void processFile(File file) {
        if (!(file.getName().endsWith(".xls") || file.getName().endsWith(".xlsx"))) {
            LOGGER.warn("Invalid file format: {}", file.getName());
            return;
        }

        String newFileName = "Processed_" + file.getName();
        File outputFile = new File(Main.UPLOAD_DIR + newFileName);

        try {
            LOGGER.info("Processing file: {}", file.getName());
            ExcelProcessor.readAndWriteExcelFile(file, outputFile);
            LOGGER.info("File processed successfully: {} at location {}", newFileName, outputFile.getAbsolutePath());
        } catch (IOException e) {
            LOGGER.error("Failed to process file: {}", file.getName(), e);
        }
    }
}
