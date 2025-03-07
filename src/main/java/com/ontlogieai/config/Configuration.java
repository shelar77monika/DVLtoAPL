package com.ontlogieai.config;

import com.fasterxml.jackson.annotation.JsonProperty;
import lombok.Getter;
import lombok.Setter;

import java.util.List;
import java.util.Map;

@Getter
@Setter
public class Configuration {

    private int deviceTagIndex;
    private int pointDescriptionIndex;
    private int pointDescriptionIndexInOutputFile;
    private int deviceTagIndexInOutputFile;

    @JsonProperty("deviceTagMapping")
    private Map<String, String> deviceTagMapping;

    @JsonProperty("deviceKeyMapping")
    private Map<String, List<String>> deviceKeyMapping;

    @JsonProperty("requiredHeaders")
    private List<String> requiredHeaders;
}
