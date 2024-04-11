package com.computation.estimate.entity;

public class Compulation {

	private String compulationNumber;
	private String compulationName;
	private double compulationValue;

	public Compulation(String compulationNumber, String compulationName,
			double compulationValue) {
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

}
