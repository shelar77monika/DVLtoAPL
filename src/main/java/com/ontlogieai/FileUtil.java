package com.ontlogieai;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;

public class FileUtil {

    private static final Logger LOGGER = LoggerFactory.getLogger(FileUtil.class);

    public static void ensureDirectoryExists(String dirPath) {
        File dir = new File(dirPath);
        if (!dir.exists() && dir.mkdirs()) {
            LOGGER.info("Directory created: {}", dirPath);
        } else if (!dir.exists()) {
            LOGGER.error("Failed to create directory: {}", dirPath);
        }
    }
}
