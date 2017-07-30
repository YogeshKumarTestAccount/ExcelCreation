/**
 * 
 */
package com.charter.excelapi;

/**
 * @author YOGESH
 *
 */
public enum OSMHeaderValues {

	CSGORDERID("CSG Order ID"), CUSTOMERID("Customer ID "), ACCOUNTID(
			"Account ID"), SCENARIO("Scenario"), SERVICETYPE("Service Type"), SERVICEID(
			"Service ID"), CMMAC("CMMAC"), MTAMAC("MTAMAC"), UIMVALIDATION(
			"UIM Validation"), PROVALIDATION("PRO Validation"), OVERALLSTATUS(
			"Overall Status");

	private String name;

	OSMHeaderValues(String name) {
		this.setName(name);
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}
}
