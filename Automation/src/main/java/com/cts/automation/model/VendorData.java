package com.cts.automation.model;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.stereotype.Component;

@Component
@ConfigurationProperties

public class VendorData {

	
	private Map<String, Map<String, String>> vendor = new HashMap<>();

    public Map<String, Map<String, String>> getVendor() {
        return vendor;
    }

    public void setMapa(Map<String, Map<String, String>> vendor) {
        this.vendor = vendor;
    }
}
