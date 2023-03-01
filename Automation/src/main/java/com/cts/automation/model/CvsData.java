package com.cts.automation.model;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.stereotype.Component;

@Component
@ConfigurationProperties

public class CvsData {

	
	private Map<String, Map<String, String>> cvs = new HashMap<>();

    public Map<String, Map<String, String>> getCvs() {
        return cvs;
    }

    public void setMapa(Map<String, Map<String, String>> cvs) {
        this.cvs = cvs;
    }
}
