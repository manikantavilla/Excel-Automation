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
	
	private Date resourseDate;
	
	private String sowAmount;
	
	private String textArea;
	
	private String empId;
	
	private boolean sow;
	
	private String sowName;
	
}
