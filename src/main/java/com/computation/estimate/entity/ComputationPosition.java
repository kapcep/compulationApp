package com.computation.estimate.entity;

public class ComputationPosition {

	private int computationPositionId;
	private Computation computation;
	private ComputationTypePosition computationTypePosition;
	private int positionNumberInComputation;
	private String positionCode;
	private String explanation;
	private String computationPositionName;
	private ComputationPositionUnitOfMeasurement computationPositionUnitOfMeasurement;
	private double amount;
	private double computationPositionPriceValue;

	public ComputationPosition(int computationPositionId,
			Computation computation,
			ComputationTypePosition computationTypePosition,
			int positionNumberInComputation, String positionCode,
			String explanation, String computationPositionName,
			ComputationPositionUnitOfMeasurement computationPositionUnitOfMeasurement,
			double amount, double computationPositionPriceValue) {
		this.computationPositionId = computationPositionId;
		this.computation = computation;
		this.computationTypePosition = computationTypePosition;
		this.positionNumberInComputation = positionNumberInComputation;
		this.positionCode = positionCode;
		this.explanation = explanation;
		this.computationPositionName = computationPositionName;
		this.computationPositionUnitOfMeasurement = computationPositionUnitOfMeasurement;
		this.amount = amount;
		this.computationPositionPriceValue = computationPositionPriceValue;
	}

	public Computation getComputation() {
		return computation;
	}

	public void setComputation(Computation computation) {
		this.computation = computation;
	}

	public ComputationTypePosition getComputationTypePosition() {
		return computationTypePosition;
	}

	public void setComputationTypePosition(
			ComputationTypePosition computationTypePosition) {
		this.computationTypePosition = computationTypePosition;
	}

	public int getPositionNumberInComputation() {
		return positionNumberInComputation;
	}

	public void setPositionNumberInComputation(
			int positionNumberInComputation) {
		this.positionNumberInComputation = positionNumberInComputation;
	}

	public String getPositionCode() {
		return positionCode;
	}

	public void setPositionCode(String positionCode) {
		this.positionCode = positionCode;
	}

	public String getExplantion() {
		return explanation;
	}

	public void setExplantion(String explanation) {
		this.explanation = explanation;
	}

	public String getComputationPositionName() {
		return computationPositionName;
	}

	public void setComputationPositionName(String computationPositionName) {
		this.computationPositionName = computationPositionName;
	}

	public ComputationPositionUnitOfMeasurement getComputationPositionUnitOfMeasurement() {
		return computationPositionUnitOfMeasurement;
	}

	public void setComputationPositionUnitOfMeasurement(
			ComputationPositionUnitOfMeasurement computationPositionUnitOfMeasurement) {
		this.computationPositionUnitOfMeasurement = computationPositionUnitOfMeasurement;
	}

	public double getAmount() {
		return amount;
	}

	public void setAmount(double amount) {
		this.amount = amount;
	}

	public double getComputationPositionPriceValue() {
		return computationPositionPriceValue;
	}

	public void setComputationPositionPriceValue(
			double computationPositionPriceValue) {
		this.computationPositionPriceValue = computationPositionPriceValue;
	}

	public int getComputationPositionId() {
		return computationPositionId;
	}

	public void setComputationPositionId(int computationPositionId) {
		this.computationPositionId = computationPositionId;
	}

	@Override
	public String toString() {
		return "|" + computationPositionId + "|" + computation + "|"
				+ computationTypePosition.getComputationTypeName() + "|"
				+ positionNumberInComputation + "|" + positionCode + "|"
				+ explanation + "|" + computationPositionName + "|"
				+ computationPositionUnitOfMeasurement
						.getComputationPositionUnitOfMeasurementName()
				+ "|" + amount + "|" + computationPositionPriceValue + "|";
	}

}
