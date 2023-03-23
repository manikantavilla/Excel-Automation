package com.cts.automation.model;

import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.stereotype.Component;

@Component
@ConfigurationProperties
public class AmendmentPath {
	
	private String amendmentPath = new String();

	public String getAmendmentPath() {
		return amendmentPath;
	}

	public void setAmendmentPath(String amendmentPath) {
		this.amendmentPath = amendmentPath;
	}


}
