package com.computation.estimate.entity;

public enum ComputationTypePosition {
	R("R"), H("H"), Y("Y"), POSITION_IN_COMPULATION(" Поз. Л.С. ");

	private final String compulationTypeName;

	private ComputationTypePosition(String compulationTypeName) {
		this.compulationTypeName = compulationTypeName;
	}

	public String getComputationTypeName() {
		return compulationTypeName;
	}

}
