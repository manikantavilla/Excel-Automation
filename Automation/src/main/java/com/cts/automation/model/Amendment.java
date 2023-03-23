package com.cts.automation.model;

import java.util.Date;
import java.util.List;

import com.fasterxml.jackson.annotation.JsonProperty;

import lombok.*;

@AllArgsConstructor
@NoArgsConstructor

@Getter
@Setter
public class Amendment {
	
private String sheetName;
	
	private String costCenter;
	
	private Date startDate;
	
	private Date endDate;
	
	private Date resourceDate;
			
	private String sowAmount;
	
	private String additionResource;
	
	@JsonProperty("vendorTeam")
	private List<String> vendorTeam;
	
	private String empId;
	
	private boolean sow;
	
	private String amendmentName;
	
}
