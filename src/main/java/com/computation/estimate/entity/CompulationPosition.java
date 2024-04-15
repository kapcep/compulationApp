package com.computation.estimate.entity;

public class CompulationPosition {

	private int compulationPositionId;
	private Compulation compulation;
	private CompulationTypePosition compulationTypePosition;
	private int positionNumberInCompulation;
	private String positionCode;
	private String explantion;
	private String compulationPositionName;
	private CompulationPositionUnitOfMeasurement compulationPositionUnitOfMeasurement;
	private double amount;
	private double compulationPositionPriceValue;

	public CompulationPosition(int compulationPositionId,
			Compulation compulation,
			CompulationTypePosition compulationTypePosition,
			int positionNumberInCompulation, String positionCode,
			String explantion, String compulationPositionName,
			CompulationPositionUnitOfMeasurement compulationPositionUnitOfMeasurement,
			double amount, double compulationPositionPriceValue) {
		this.compulationPositionId = compulationPositionId;
		this.compulation = compulation;
		this.compulationTypePosition = compulationTypePosition;
		this.positionNumberInCompulation = positionNumberInCompulation;
		this.positionCode = positionCode;
		this.explantion = explantion;
		this.compulationPositionName = compulationPositionName;
		this.compulationPositionUnitOfMeasurement = compulationPositionUnitOfMeasurement;
		this.amount = amount;
		this.compulationPositionPriceValue = compulationPositionPriceValue;
	}

	public Compulation getCompulation() {
		return compulation;
	}

	public void setCompulation(Compulation compulation) {
		this.compulation = compulation;
	}

	public CompulationTypePosition getCompulationTypePosition() {
		return compulationTypePosition;
	}

	public void setCompulationTypePosition(
			CompulationTypePosition compulationTypePosition) {
		this.compulationTypePosition = compulationTypePosition;
	}

	public int getPositionNumberInCompulation() {
		return positionNumberInCompulation;
	}

	public void setPositionNumberInCompulation(
			int positionNumberInCompulation) {
		this.positionNumberInCompulation = positionNumberInCompulation;
	}

	public String getPositionCode() {
		return positionCode;
	}

	public void setPositionCode(String positionCode) {
		this.positionCode = positionCode;
	}

	public String getExplantion() {
		return explantion;
	}

	public void setExplantion(String explantion) {
		this.explantion = explantion;
	}

	public String getCompulationPositionName() {
		return compulationPositionName;
	}

	public void setCompulationPositionName(String compulationPositionName) {
		this.compulationPositionName = compulationPositionName;
	}

	public CompulationPositionUnitOfMeasurement getCompulationPositionUnitOfMeasurement() {
		return compulationPositionUnitOfMeasurement;
	}

	public void setCompulationPositionUnitOfMeasurement(
			CompulationPositionUnitOfMeasurement compulationPositionUnitOfMeasurement) {
		this.compulationPositionUnitOfMeasurement = compulationPositionUnitOfMeasurement;
	}

	public double getAmount() {
		return amount;
	}

	public void setAmount(double amount) {
		this.amount = amount;
	}

	public double getCompulationPositionPriceValue() {
		return compulationPositionPriceValue;
	}

	public void setCompulationPositionPriceValue(
			double compulationPositionPriceValue) {
		this.compulationPositionPriceValue = compulationPositionPriceValue;
	}

	public int getCompulationPositionId() {
		return compulationPositionId;
	}

	public void setCompulationPositionId(int compulationPositionId) {
		this.compulationPositionId = compulationPositionId;
	}

}
