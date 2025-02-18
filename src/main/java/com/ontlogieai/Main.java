package com.ontlogieai;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class Main {
    public static final String UPLOAD_DIR = "uploads/";
    private static final Logger logger = LoggerFactory.getLogger(Main.class);

    public static void main(String[] args) {
        logger.info("Application starting...");
        FileUtil.ensureDirectoryExists(UPLOAD_DIR);
        UIUtil.setLookAndFeel();
        UIUtil.createAndShowGUI();
        logger.info("Application started successfully.");
    }

}
