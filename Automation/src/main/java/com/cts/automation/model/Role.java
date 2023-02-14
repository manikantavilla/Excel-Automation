package com.cts.automation.model;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@AllArgsConstructor
@NoArgsConstructor

@Getter
@Setter
public class Role {
	
	private String role;
	
	private String location;
	
	private int rate;
	
	private int total;

}
