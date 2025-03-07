package com.ontlogieai.transformation;

import java.util.Set;

public class DeviceTagUtility {

    private static final Set<String> DEVICE_TYPES = Set.of(
            "TT", "FT", "MT", "PT", "ACU", "XC", "TC", "QIT", "UPS",
            "VAV", "XT", "XA", "FCV", "KS", "XI", "PMP"
    );

    public static String getDeviceType(String deviceTag) {
        if (deviceTag == null || deviceTag.isEmpty()) return "";

        return DEVICE_TYPES.stream()
                .filter(deviceTag::contains)
                .findFirst()
                .orElse("");
    }
}
