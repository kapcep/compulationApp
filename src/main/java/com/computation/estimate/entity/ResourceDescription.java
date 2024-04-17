package com.computation.estimate.entity;

public class ResourceDescription {

	private int resourceDescriptionId;
	private ComputationPosition computationPosition;
	private MarkOfResourceDescrition markOfResourceDescrition;
	private String resourceCode;
	private double computationResourcePrice;
	private double standardConsumptionOfTheResource;
	private String resourceDescriptionName;
	private ResourceDescriptionUnitOfMeasurement resourceDescriptionUnitOfMeasurement;

	public ResourceDescription(int resourceDescriptionId,
			ComputationPosition computationPosition,
			MarkOfResourceDescrition markOfResourceDescrition,
			String resourceCode, double computationResourcePrice,
			double standardConsumptionOfTheResource,
			String resourceDescriptionName,
			ResourceDescriptionUnitOfMeasurement resourceDescriptionUnitOfMeasurement) {
		this.resourceDescriptionId = resourceDescriptionId;
		this.computationPosition = computationPosition;
		this.markOfResourceDescrition = markOfResourceDescrition;
		this.resourceCode = resourceCode;
		this.computationResourcePrice = computationResourcePrice;
		this.standardConsumptionOfTheResource = standardConsumptionOfTheResource;
		this.resourceDescriptionName = resourceDescriptionName;
		this.resourceDescriptionUnitOfMeasurement = resourceDescriptionUnitOfMeasurement;
	}

	public int getResourceDescriptionId() {
		return resourceDescriptionId;
	}

	public void setResourceDescriptionId(int resourceDescriptionId) {
		this.resourceDescriptionId = resourceDescriptionId;
	}

	public ComputationPosition getComputationPosition() {
		return computationPosition;
	}

	public void setComputationPosition(
			ComputationPosition computationPosition) {
		this.computationPosition = computationPosition;
	}

	public MarkOfResourceDescrition getMarkOfResourceDescrition() {
		return markOfResourceDescrition;
	}

	public void setMarkOfResourceDescrition(
			MarkOfResourceDescrition markOfResourceDescrition) {
		this.markOfResourceDescrition = markOfResourceDescrition;
	}

	public String getResourceCode() {
		return resourceCode;
	}

	public void setResourceCode(String resourceCode) {
		this.resourceCode = resourceCode;
	}

	public double getComputationResourcePrice() {
		return computationResourcePrice;
	}

	public void setComputationResourcePrice(double computationResourcePrice) {
		this.computationResourcePrice = computationResourcePrice;
	}

	public double getStandardConsumptionOfTheResource() {
		return standardConsumptionOfTheResource;
	}

	public void setStandardConsumptionOfTheResource(
			double standardConsumptionOfTheResource) {
		this.standardConsumptionOfTheResource = standardConsumptionOfTheResource;
	}

	public String getResourceDescriptionName() {
		return resourceDescriptionName;
	}

	public void setResourceDescriptionName(String resourceDescriptionName) {
		this.resourceDescriptionName = resourceDescriptionName;
	}

	public ResourceDescriptionUnitOfMeasurement getResourceDescriptionUnitOfMeasurement() {
		return resourceDescriptionUnitOfMeasurement;
	}

	public void setResourceDescriptionUnitOfMeasurement(
			ResourceDescriptionUnitOfMeasurement resourceDescriptionUnitOfMeasurement) {
		this.resourceDescriptionUnitOfMeasurement = resourceDescriptionUnitOfMeasurement;
	}

	@Override
	public String toString() {
		return "|" + resourceDescriptionId + "|"
				+ computationPosition.getComputationPositionName() + "|"
				+ markOfResourceDescrition + "|" + resourceCode + "|"
				+ resourceDescriptionName + "|"
				+ resourceDescriptionUnitOfMeasurement
						.getResourceDescriptionUnitOfMeasurementName()
				+ "|" + standardConsumptionOfTheResource + "|"
				+ computationResourcePrice + "|";
	}

}
