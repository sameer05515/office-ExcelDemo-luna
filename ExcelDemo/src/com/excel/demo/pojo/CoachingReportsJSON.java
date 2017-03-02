package com.excel.demo.pojo;

import java.util.ArrayList;
import java.util.List;

/*
 [
	{
		coachName: <>
		totalForCoach:<>
		totalForOthers{selected role}:<>
		total:<>
		
		reportList:[
			{
				actionName:<>
				percentage:<>
			},
			{
				actionName:<>
				percentage:<>
			}
		
		]
		
		gridList:[
		
		CoachingSessionsJson
		CoachingSessionsJson
		CoachingSessionsJson
		]

	}





]
 */
public class CoachingReportsJSON {
	private String coachName="Premandra Kumar";
	private int totalForCoach=20;
	private String selectedRoleName="Quality Managers";
	private int totalForOthers=180;
	private String totalCaption="TOTAL";
	private int total=200;
	
	private List<Reports> reportsList=new ArrayList<Reports>();

	public String getCoachName() {
		return coachName;
	}

	public void setCoachName(String coachName) {
		this.coachName = coachName;
	}

	public int getTotalForCoach() {
		return totalForCoach;
	}

	public void setTotalForCoach(int totalForCoach) {
		this.totalForCoach = totalForCoach;
	}

	public int getTotalForOthers() {
		return totalForOthers;
	}

	public void setTotalForOthers(int totalForOthers) {
		this.totalForOthers = totalForOthers;
	}

	public int getTotal() {
		return total;
	}

	public void setTotal(int total) {
		this.total = total;
	}

	public List<Reports> getReportsList() {
		return reportsList;
	}

	public void setReportsList(List<Reports> reportsList) {
		this.reportsList = reportsList;
	}

	public String getSelectedRoleName() {
		return selectedRoleName;
	}

	public void setSelectedRoleName(String selectedRoleName) {
		this.selectedRoleName = selectedRoleName;
	}

	public String getTotalCaption() {
		return totalCaption;
	}

	public void setTotalCaption(String totalCaption) {
		this.totalCaption = totalCaption;
	}
	
	
}

