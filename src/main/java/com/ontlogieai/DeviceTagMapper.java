package com.ontlogieai;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.HashMap;
import java.util.Map;
import java.util.Optional;

public class DeviceTagMapper {

    private static final Logger LOGGER = LoggerFactory.getLogger(DeviceTagMapper.class.getName());
    private static final Map<String, String> STANDARD_DEVICE_TAGS = new HashMap<>();

    static {
        // Populate the standard device tag mapping
        //TT
        STANDARD_DEVICE_TAGS.put("TT-Potable Water - Temperature", "J100-06-2TT-001");
        STANDARD_DEVICE_TAGS.put("TT-Potable  Hot Water", "J100-06-2TT-002");
        STANDARD_DEVICE_TAGS.put("TT-Non Potable Water","J130-06-2TT-001");
        STANDARD_DEVICE_TAGS.put("TT-Chilled Water - Supply Temperature", "J460-01-2TT-612");
        STANDARD_DEVICE_TAGS.put("TT-Chilled Water - Return Temperature", "J460-01-2TT-613");
        STANDARD_DEVICE_TAGS.put("TT-Supply Air", "J460-01-2TT-614");
        STANDARD_DEVICE_TAGS.put("TT-Return Air", "J460-01-2TT-616");

        //FT
        STANDARD_DEVICE_TAGS.put("FT-Chilled Water", "J460-01-2FT-601");
        STANDARD_DEVICE_TAGS.put("FT-Hot Water", "J460-01-2FT-602");
        STANDARD_DEVICE_TAGS.put("FT-Supply Air Flow", "J460-01-2FT-603");
        STANDARD_DEVICE_TAGS.put("FT-Return Air Flow", "J460-01-2FT-604");
        STANDARD_DEVICE_TAGS.put("FT-Potable Water", "J100-06-2FT-001");
        STANDARD_DEVICE_TAGS.put("FT-Non Potable Water", "J130-06-2FT-001");
        STANDARD_DEVICE_TAGS.put("FT-Compressed Air", "J305-06-2FT-001");
        STANDARD_DEVICE_TAGS.put("FT-Carbon Dioxide Gas","J305-06-2FT-001");
        STANDARD_DEVICE_TAGS.put("FT-Nitrogen Gas","J330-06-2FT-001");
        STANDARD_DEVICE_TAGS.put("FT-Demi Water","J140-06-2FT-001");

        //MT
        STANDARD_DEVICE_TAGS.put("MT-Supply Air Humidity", "J460-01-2MT-614");
        STANDARD_DEVICE_TAGS.put("MT-Compressed Air", "J305-06-2MT-001");
        STANDARD_DEVICE_TAGS.put("MT-Compressed Air - Dewpoint", "J305-06-2MT-001");
        STANDARD_DEVICE_TAGS.put("MT-Humidity", "J460-01-2MT-601");
        STANDARD_DEVICE_TAGS.put("MT-Chilled Water Valve - Controller", "J460-01-2FCV-643");

        //PT
        STANDARD_DEVICE_TAGS.put("PT-Non Potable Water - Pressure", "J130-06-2PT-001");
        STANDARD_DEVICE_TAGS.put("PT-Compressed Air - Pressure", "J305-06-2PT-001");
        STANDARD_DEVICE_TAGS.put("PT-Carbon Dioxide Gas - Pressure", "J320-06-2PT-001");
        STANDARD_DEVICE_TAGS.put("PT-Nitrogen Gas - Pressure", "J330-06-2PT-001");
        STANDARD_DEVICE_TAGS.put("PT-Demi Water - Inlet Pressure", "J140-06-2PT-001");
        STANDARD_DEVICE_TAGS.put("PT-Demi Water - Outlet Pressure", "J140-06-2PT-002");
        STANDARD_DEVICE_TAGS.put("PT-Demi Water - Return Pressure", "J140-06-2PT-003");
        STANDARD_DEVICE_TAGS.put("PT-Pressure", "J460-02-2PT-902");

        //ACU
        STANDARD_DEVICE_TAGS.put("ACU-Fan Speed", "J460-02-1ACU-601");
        STANDARD_DEVICE_TAGS.put("ACU-Fan Coil Unit Control", "J460-01-1ACU-618");

        //XC
        STANDARD_DEVICE_TAGS.put("XC-Exhaust Fan", "J460-02-2B-902");

        //XT
        STANDARD_DEVICE_TAGS.put("XT-Occupied", "J460-01-2XT-002");
        STANDARD_DEVICE_TAGS.put("XT-CO2 Concentration", "J460-01-2XT-001");

        //XA
        STANDARD_DEVICE_TAGS.put("XA-Thermal Fault Signal", "J270-06-2XA-001");
        STANDARD_DEVICE_TAGS.put("XA-Surge Voltage Arrester Signal", "J270-06-2XA-004");
        STANDARD_DEVICE_TAGS.put("XA-Common Fire Alarm", "J270-06-2XA-005");
        STANDARD_DEVICE_TAGS.put("XA-Circuit Breaker Tripped", "J229-06-2XA-101");
        STANDARD_DEVICE_TAGS.put("XA-Voltage Surge Arrestor", "J229-06-2XA-102");
        STANDARD_DEVICE_TAGS.put("XA-UPS Alarm", "J460-01-2XA-802");

        //FCV
        STANDARD_DEVICE_TAGS.put("FCV-Reheater Valve Control", "J460-02-2FCV-623");
        STANDARD_DEVICE_TAGS.put("FCV-Heating Valve Control", "J460-01-2FCV-002");
        STANDARD_DEVICE_TAGS.put("FCV-Cooling Valve Control", "J460-01-2FCV-006");
        STANDARD_DEVICE_TAGS.put("FCV-Chilled Water Valve", "J460-01-2FCV-643");

        //KS
        STANDARD_DEVICE_TAGS.put("KS-Labs Day Extension Timer - Timer","J460-02-2KS-603");

        //XI
        STANDARD_DEVICE_TAGS.put("XI-Labs Day Extension Timer - Indicator","J460-02-2XI-603");

        //PMP
        STANDARD_DEVICE_TAGS.put("PMP-Chilled Water Circulation Pump","J460-01-1PMP-601");

        //QIT
        STANDARD_DEVICE_TAGS.put("QIT-Energy Meter", "J229-06-1QIT-001");

        //UPS
        STANDARD_DEVICE_TAGS.put("UPS-UPS", "J232-06-1UPS-001");

        //VAV
        STANDARD_DEVICE_TAGS.put("VAV-Return Air Flow Control", "J460-01-1VAV-001");
        STANDARD_DEVICE_TAGS.put("VAV-Supply Air Flow Control", "J460-01-1VAV-603");
        STANDARD_DEVICE_TAGS.put("VAV-Fume hood", "J460-02-1VAV-606");
        STANDARD_DEVICE_TAGS.put("VAV-Air Fow Control", "J460-02-1VAV-607");

        //TC
        STANDARD_DEVICE_TAGS.put("TC-Room Controller","J460-01-2TC-601");

        //XCV
        STANDARD_DEVICE_TAGS.put("XCV-Legionella Dump Valve", "J100-06-2XCV-001");

    }

    public static String getStandardDeviceTag(String deviceTag, String pointDescription) {
        if ("J460-02-2TT-717".equalsIgnoreCase(deviceTag)) {
            LOGGER.info("Here we need to debug...");
        }

        String keyPrefix = getDeviceType(deviceTag);
        String deviceKeyPostfix = getDeviceKey(keyPrefix, pointDescription);
        String deviceKey = keyPrefix + "-" + deviceKeyPostfix;

        return STANDARD_DEVICE_TAGS.getOrDefault(deviceKey, "");
    }

    private static String getDeviceKey(String keyPrefix, String pointDescriptor) {
        if (pointDescriptor == null) return "";

        Map<String, String[]> deviceMappings = new HashMap<>(Map.of(
                "TT", new String[]{"Potable Water - Temperature", "Potable  Hot Water", "Non Potable Water", "Chilled Water - Supply Temperature", "Chilled Water - Return Temperature", "Supply Air", "Return Air"},
                "FT", new String[]{"Chilled Water", "Hot Water", "Supply Air Flow", "Return Air Flow", "Potable Water", "Compressed Air", "Carbon Dioxide Gas", "Nitrogen Gas", "Demi Water"},
                "MT", new String[]{"Supply Air Humidity", "Humidity", "Chilled Water Valve - Controller", "Compressed Air - Dewpoint", "Compressed Air"},
                "PT", new String[]{"Non Potable Water - Pressure", "Compressed Air - Pressure", "Carbon Dioxide Gas - Pressure", "Nitrogen Gas - Pressure", "Demi Water - Inlet Pressure", "Demi Water - Outlet Pressure", "Demi Water - Return Pressure", "Pressure"},
                "ACU", new String[]{"Fan Speed", "Fan Coil Unit Control"},
                "XC", new String[]{"Exhaust Fan"},
                "QIT", new String[]{"Energy Meter"},
                "UPS", new String[]{"UPS"},
                "VAV", new String[]{"Return Air Flow Control", "Supply Air Flow Control", "Fume hood", "Air Flow Control"},
                "TC", new String[]{"Room Controller"}

        ));
        deviceMappings.put("XT", new String[]{"Occupied", "CO2 Concentration"});
        deviceMappings.put("XA", new String[]{ "Thermal Fault Signal", "Surge Voltage Arrester Signal", "Common Fire Alarm", "Circuit Breaker Tripped", "Voltage Surge Arrestor", "UPS Alarm" });
        deviceMappings.put("FCV", new String[]{ "Reheater Valve Control", "Heating Valve Control", "Cooling Valve Control", "Chilled Water Valve" });
        deviceMappings.put("KS", new String[]{ "Labs Day Extension Timer - Timer" });
        deviceMappings.put("XI", new String[]{ "Labs Day Extension Timer - Indicator" });
        deviceMappings.put("PMP", new String[]{ "Chilled Water Circulation Pump" });


        return Optional.ofNullable(deviceMappings.get(keyPrefix))
                .flatMap(keys -> findMatchingKey(keys, pointDescriptor))
                .orElse("");
    }

    private static Optional<String> findMatchingKey(String[] keys, String descriptor) {
        for (String key : keys) {
            if (descriptor.contains(key)) {
                return Optional.of(key);
            }
        }
        return Optional.empty();
    }

    private static String getDeviceType(String deviceTag) {
        if (deviceTag == null) return "";

        Map<String, String> deviceTypeMappings = new HashMap<>(Map.of(
                "TT", "TT",
                "FT", "FT",
                "MT", "MT",
                "PT", "PT",
                "ACU", "ACU",
                "XC", "XC",
                "TC", "TC",
                "QIT", "QIT",
                "UPS", "UPS",
                "VAV", "VAV"
        ));
        deviceTypeMappings.put("TC","TC");
        deviceTypeMappings.put("XT", "XT");
        deviceTypeMappings.put("XA", "XA");
        deviceTypeMappings.put("FCV", "FCV");
        deviceTypeMappings.put("KS", "KS");
        deviceTypeMappings.put("XI", "XI");
        deviceTypeMappings.put("PMP", "PMP");

        return deviceTypeMappings.entrySet().stream()
                .filter(entry -> deviceTag.contains(entry.getKey()))
                .map(Map.Entry::getValue)
                .findFirst()
                .orElse("");
    }

}
