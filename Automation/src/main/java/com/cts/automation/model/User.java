package com.cts.automation.model;

import java.util.Date;

import lombok.*;

@AllArgsConstructor
@NoArgsConstructor

@Getter
@Setter
public class User {
	
	private String costcenter;
	
	private Date startdate;
	
	private Date enddate;
	
	private String vendorprojectteammember1;
	
	private String vendorprojectteammember2;
	
	private String vendorprojectteammember3;
	
	private String cvsprojectteammember1;
	
//	private String deliverables;

	
}
