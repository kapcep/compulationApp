package com.computation.estimate.entity;

public class Computation {

	private int сomputationId;
	private String сomputationNumber;
	private String сomputationName;
	private double сomputationValue;

	public Computation(int сomputationId, String сomputationNumber,
			String сomputationName, double сomputationValue) {
		this.сomputationId = сomputationId;
		this.сomputationNumber = сomputationNumber;
		this.сomputationName = сomputationName;
		this.сomputationValue = сomputationValue;
	}

	public String getComputationNumber() {
		return сomputationNumber;
	}

	public void setComputationNumber(String сomputationNumber) {
		this.сomputationNumber = сomputationNumber;
	}

	public String getComputationName() {
		return сomputationName;
	}

	public void setComputationName(String сomputationName) {
		this.сomputationName = сomputationName;
	}

	public double getComputationValue() {
		return сomputationValue;
	}

	public void setComputationValue(double сomputationValue) {
		this.сomputationValue = сomputationValue;
	}

	public int getComputationId() {
		return сomputationId;
	}

	public void setComputationId(int сomputationId) {
		this.сomputationId = сomputationId;
	}

}
