package com.computation.estimate.entity;

public class ResourceDescription {

	private int resourceDescriptionId;
	private CompulationPosition compulationPosition;
	private MarkOfResourceDescrition markOfResourceDescrition;
	private String resourceCode;
	private double compulationResourcePrice;
	private double standardConsumptionOfTheResource;
	private String ResourceDescriptionName;
	private String resourceDescriptionUnitOfMeasurement;

	public ResourceDescription(int resourceDescriptionId,
			CompulationPosition compulationPosition,
			MarkOfResourceDescrition markOfResourceDescrition,
			String resourceCode, double compulationResourcePrice,
			double standardConsumptionOfTheResource,
			String resourceDescriptionName,
			String resourceDescriptionUnitOfMeasurement) {
		this.resourceDescriptionId = resourceDescriptionId;
		this.compulationPosition = compulationPosition;
		this.markOfResourceDescrition = markOfResourceDescrition;
		this.resourceCode = resourceCode;
		this.compulationResourcePrice = compulationResourcePrice;
		this.standardConsumptionOfTheResource = standardConsumptionOfTheResource;
		ResourceDescriptionName = resourceDescriptionName;
		this.resourceDescriptionUnitOfMeasurement = resourceDescriptionUnitOfMeasurement;
	}

	public CompulationPosition getCompulationPosition() {
		return compulationPosition;
	}

	public void setCompulationPosition(
			CompulationPosition compulationPosition) {
		this.compulationPosition = compulationPosition;
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

	public double getCompulationResourcePrice() {
		return compulationResourcePrice;
	}

	public void setCompulationResourcePrice(double compulationResourcePrice) {
		this.compulationResourcePrice = compulationResourcePrice;
	}

	public double getStandardConsumptionOfTheResource() {
		return standardConsumptionOfTheResource;
	}

	public void setStandardConsumptionOfTheResource(
			double standardConsumptionOfTheResource) {
		this.standardConsumptionOfTheResource = standardConsumptionOfTheResource;
	}

	public String getResourceDescriptionName() {
		return ResourceDescriptionName;
	}

	public void setResourceDescriptionName(String resourceDescriptionName) {
		ResourceDescriptionName = resourceDescriptionName;
	}

	public String getUnitOfMeasurement() {
		return resourceDescriptionUnitOfMeasurement;
	}

	public void setUnitOfMeasurement(String unitOfMeasurement) {
		this.resourceDescriptionUnitOfMeasurement = unitOfMeasurement;
	}

	public int getResourceDescriptionId() {
		return resourceDescriptionId;
	}

	public void setResourceDescriptionId(int resourceDescriptionId) {
		this.resourceDescriptionId = resourceDescriptionId;
	}

	public String getResourceDescriptionUnitOfMeasurement() {
		return resourceDescriptionUnitOfMeasurement;
	}

	public void setResourceDescriptionUnitOfMeasurement(
			String resourceDescriptionUnitOfMeasurement) {
		this.resourceDescriptionUnitOfMeasurement = resourceDescriptionUnitOfMeasurement;
	}

}
