package com.computation.estimate.entity;

public class CompulationPositionUnitOfMeasurement {

	private String compulationPositionUnitOfMeasurementName;

	public CompulationPositionUnitOfMeasurement(
			String compulationPositionUnitOfMeasurementName) {
		this.compulationPositionUnitOfMeasurementName = compulationPositionUnitOfMeasurementName;
	}

	public String getCompulationPositionUnitOfMeasurementName() {
		return compulationPositionUnitOfMeasurementName;
	}

	public void setCompulationPositionUnitOfMeasurementName(
			String compulationPositionUnitOfMeasurementName) {
		this.compulationPositionUnitOfMeasurementName = compulationPositionUnitOfMeasurementName;
	}

	@Override
	public String toString() {
		return "CompulationPositionUnitOfMeasurement [compulationPositionUnitOfMeasurementName="
				+ compulationPositionUnitOfMeasurementName + "]";
	}

}
