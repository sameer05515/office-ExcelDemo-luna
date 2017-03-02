package com.excel.demo.pojo;
public class Reports{
	private String actionName="";
	public Reports(String actionName, double percentage) {
		super();
		this.actionName = actionName;
		this.percentage = percentage;
	}
	private double percentage=0.0;
	public String getActionName() {
		return actionName;
	}
	public void setActionName(String actionName) {
		this.actionName = actionName;
	}
	public double getPercentage() {
		return percentage;
	}
	public void setPercentage(double percentage) {
		this.percentage = percentage;
	}
	
	
}
