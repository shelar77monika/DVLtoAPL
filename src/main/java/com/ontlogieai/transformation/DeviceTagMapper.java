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

    private Configuration configuration;

    public DeviceTagMapper(){
        configuration = ConfigurationReader.getConfiguration();
    }

    public String getStandardDeviceTag(String deviceTag, String pointDescription) {
        if ("J460-02-2TT-717".equalsIgnoreCase(deviceTag)) {
            LOGGER.info("Here we need to debug...");
        }

        String keyPrefix = DeviceTagUtility.getDeviceType(deviceTag);
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

}
