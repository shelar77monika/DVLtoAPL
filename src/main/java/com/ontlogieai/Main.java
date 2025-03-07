package com.ontlogieai;

import com.ontlogieai.file.FileUtil;
import com.ontlogieai.ui.UIUtil;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;


public class Main {
    public static final String UPLOAD_DIR = "uploads/";
    private static final Logger logger = LoggerFactory.getLogger(Main.class);

    public static void main(String[] args) {
        logger.info("Application starting...");
        FileUtil.ensureDirectoryExists(UPLOAD_DIR);
        UIUtil uiUtil = new UIUtil();
        uiUtil.setLookAndFeel();
        uiUtil.createAndShowGUI();
        logger.info("Application started successfully.");
        System.out.println("Application is started");
    }

}
