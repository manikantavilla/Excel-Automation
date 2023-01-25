package com.cts.automation.model;

import lombok.*;

@AllArgsConstructor
@NoArgsConstructor

@Getter
@Setter
public class User {

	private String month;
	
	private String year;
	
	private String startdate;
	
	private String enddate;
	
	private int totalamount;
	
	private String vendorprojectteammember1;
	
	private String vendorprojectteammember2;
	
	private String vendorprojectteammember3;
	
	private String cvsprojectteammember1;
	
	private int deliverables;
}
