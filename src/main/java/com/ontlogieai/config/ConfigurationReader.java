package com.ontlogieai.config;

import com.fasterxml.jackson.databind.ObjectMapper;

import java.io.IOException;
import java.io.InputStream;

public class ConfigurationReader {

    private static final String CONFIG_FILE = "config.json"; // Change as per your file location

    private static Configuration configuration;

    public static Configuration loadConfig() {
        ObjectMapper objectMapper = new ObjectMapper();
        try (InputStream inputStream = ConfigurationReader.class.getClassLoader().getResourceAsStream(CONFIG_FILE)) {
            if (inputStream == null) {
                throw new RuntimeException("Configuration file not found: " + CONFIG_FILE);
            }
            return objectMapper.readValue(inputStream, Configuration.class);
        } catch (IOException e) {
            throw new RuntimeException("Error loading configuration", e);
        }
    }

    public static Configuration getConfiguration(){
        if(null == configuration){
            configuration = loadConfig();
            return configuration;
        }
        return configuration;
    }
}
