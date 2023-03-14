package com.cts.automation.model;

import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.stereotype.Component;

@Component
@ConfigurationProperties
public class SubmitPath {
	
	private String submitPath = new String();

	public String getSubmitPath() {
		return submitPath;
	}

	public void setSubmitPath(String submitPath) {
		this.submitPath = submitPath;
	}


}
