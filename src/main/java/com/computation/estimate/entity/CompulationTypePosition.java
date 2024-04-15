package com.computation.estimate.entity;

public enum CompulationTypePosition {
	R("R"), H("H"), Y("Y"), POSITION_IN_COMPULATION("Поз. Л.С.");

	private final String compulationTypeName;

	private CompulationTypePosition(String compulationTypeName) {
		this.compulationTypeName = compulationTypeName;
	}

	public String getCompulationTypeName() {
		return compulationTypeName;
	}

}
