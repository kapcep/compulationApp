package com.computation.estimate.entity;

public class Compulation {

	private int compulationId;
	private String compulationNumber;
	private String compulationName;
	private double compulationValue;

	public Compulation(int compulationId, String compulationNumber,
			String compulationName, double compulationValue) {
		this.compulationId = compulationId;
		this.compulationNumber = compulationNumber;
		this.compulationName = compulationName;
		this.compulationValue = compulationValue;
	}

	public String getCompulationNumber() {
		return compulationNumber;
	}

	public void setCompulationNumber(String compulationNumber) {
		this.compulationNumber = compulationNumber;
	}

	public String getCompulationName() {
		return compulationName;
	}

	public void setCompulationName(String compulationName) {
		this.compulationName = compulationName;
	}

	public double getCompulationValue() {
		return compulationValue;
	}

	public void setCompulationValue(double compulationValue) {
		this.compulationValue = compulationValue;
	}

	public int getCompulationId() {
		return compulationId;
	}

	public void setCompulationId(int compulationId) {
		this.compulationId = compulationId;
	}

}
