package com.cts.automation.model;

import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.stereotype.Component;

@Component
@ConfigurationProperties
public class SOWPath {
	
	private String sowPath = new String();

	public String getSowPath() {
		return sowPath;
	}

	public void setSowPath(String sowPath) {
		this.sowPath = sowPath;
	}
}
