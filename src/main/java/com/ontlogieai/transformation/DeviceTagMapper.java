package com.ontlogieai.transformation;

import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.ontlogieai.config.Configuration;
import com.ontlogieai.config.ConfigurationReader;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.InputStream;
import java.util.*;

public class DeviceTagMapper {

    private static final Logger LOGGER = LoggerFactory.getLogger(DeviceTagMapper.class.getName());
    private static final Map<String, String> STANDARD_DEVICE_TAGS = new HashMap<>();
    private static final Map<String, List<String>> DEVICE_KEY_MAPPINGS = new HashMap<>();
    private static final Set<String> DEVICE_TYPES = Set.of(
            "TT", "FT", "MT", "PT", "ACU", "XC", "TC", "QIT", "UPS",
            "VAV", "XT", "XA", "FCV", "KS", "XI", "PMP"
    );

    private Configuration configuration;

    public DeviceTagMapper(){
        configuration = ConfigurationReader.getConfiguration();
    }

    static {
        // Populate the standard device tag mapping
        loadDeviceTags();
        loadDeviceKeys();

    }

    private static void loadDeviceTags() {
        ObjectMapper objectMapper = new ObjectMapper();
        String resourcePath = "config.json"; // Can be changed to an external path

        try (InputStream inputStream = DeviceTagMapper.class.getClassLoader().getResourceAsStream(resourcePath)) {
            if (inputStream != null) {
                Map<String, Object> rootMap = objectMapper.readValue(inputStream, new TypeReference<>() {});
                Map<String, String> deviceTags = (Map<String, String>) rootMap.get("deviceTagMapping"); // Extract the nested map

                if (deviceTags != null) {
                    STANDARD_DEVICE_TAGS.putAll(deviceTags);
                } else {
                    LOGGER.warn("Key 'deviceTagMapping' not found in JSON.");
                }
            } else {
                LOGGER.warn("Resource '{}' not found. Ensure the file is present in the resources folder.", resourcePath);
            }
        } catch (IOException e) {
            LOGGER.error("Error loading device tags from JSON", e);
        }
    }

    private static void loadDeviceKeys() {
        ObjectMapper objectMapper = new ObjectMapper();
        String resourcePath = "config.json"; // Can be changed to an external path

        try (InputStream inputStream = DeviceTagMapper.class.getClassLoader().getResourceAsStream(resourcePath)) {
            if (inputStream != null) {
                Map<String, Object> rootMap = objectMapper.readValue(inputStream, new TypeReference<>() {});
                Map<String, List<String>> deviceKeyMappings = (Map<String, List<String>>) rootMap.get("deviceKeyMapping"); // Extract the nested map

                if (deviceKeyMappings != null) {
                    DEVICE_KEY_MAPPINGS.putAll(deviceKeyMappings);
                } else {
                    LOGGER.warn("Key 'devicekeyMapping' not found in JSON.");
                }
            } else {
                LOGGER.warn("Resource '{}' not found. Ensure the file is present in the resources folder.", resourcePath);
            }
        } catch (IOException e) {
            LOGGER.error("Error loading device tags from JSON", e);
        }
    }


    public String getStandardDeviceTag(String deviceTag, String pointDescription) {
        if ("J460-02-2TT-717".equalsIgnoreCase(deviceTag)) {
            LOGGER.info("Here we need to debug...");
        }

        String keyPrefix = getDeviceType(deviceTag);
        String deviceKeyPostfix = getDeviceKey(keyPrefix, pointDescription);
        String deviceKey = keyPrefix + "-" + deviceKeyPostfix;

        return configuration.getDeviceTagMapping().getOrDefault(deviceKey, "");
    }

    private String getDeviceKey(String keyPrefix, String pointDescriptor) {
        if (pointDescriptor == null) return "";

        return Optional.ofNullable(configuration.getDeviceKeyMapping().get(keyPrefix))
                .flatMap(keys -> findMatchingKey(keys, pointDescriptor))
                .orElse("");
    }

    private static Optional<String> findMatchingKey(List<String> keys, String descriptor) {
        for (String key : keys) {
            if (descriptor.contains(key)) {
                return Optional.of(key);
            }
        }
        return Optional.empty();
    }

    private static String getDeviceType(String deviceTag) {
        if (deviceTag == null || deviceTag.isEmpty()) return "";

        return DEVICE_TYPES.stream()
                .filter(deviceTag::contains)
                .findFirst()
                .orElse("");
    }

}
