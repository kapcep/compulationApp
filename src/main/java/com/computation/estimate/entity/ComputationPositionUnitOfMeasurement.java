package com.computation.estimate.entity;

public class ComputationPositionUnitOfMeasurement {

	private String сomputationPositionUnitOfMeasurementName;

	public ComputationPositionUnitOfMeasurement(
			String сomputationPositionUnitOfMeasurementName) {
		this.сomputationPositionUnitOfMeasurementName = сomputationPositionUnitOfMeasurementName;
	}

	public String getComputationPositionUnitOfMeasurementName() {
		return сomputationPositionUnitOfMeasurementName;
	}

	public void setComputationPositionUnitOfMeasurementName(
			String сomputationPositionUnitOfMeasurementName) {
		this.сomputationPositionUnitOfMeasurementName = сomputationPositionUnitOfMeasurementName;
	}

	@Override
	public String toString() {
		return "ComputationPositionUnitOfMeasurement [сomputationPositionUnitOfMeasurementName="
				+ сomputationPositionUnitOfMeasurementName + "]";
	}

}
