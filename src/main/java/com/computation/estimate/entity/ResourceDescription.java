package com.computation.estimate.entity;

public class ResourceDescription {

	private int resourceDescriptionId;
	private ComputationPosition computationPosition;
	private MarkOfResourceDescrition markOfResourceDescrition;
	private String resourceCode;
	private double computationResourcePrice;
	private double standardConsumptionOfTheResource;
	private String ResourceDescriptionName;
	private String resourceDescriptionUnitOfMeasurement;

	public ResourceDescription(int resourceDescriptionId,
			ComputationPosition computationPosition,
			MarkOfResourceDescrition markOfResourceDescrition,
			String resourceCode, double computationResourcePrice,
			double standardConsumptionOfTheResource,
			String resourceDescriptionName,
			String resourceDescriptionUnitOfMeasurement) {
		this.resourceDescriptionId = resourceDescriptionId;
		this.computationPosition = computationPosition;
		this.markOfResourceDescrition = markOfResourceDescrition;
		this.resourceCode = resourceCode;
		this.computationResourcePrice = computationResourcePrice;
		this.standardConsumptionOfTheResource = standardConsumptionOfTheResource;
		ResourceDescriptionName = resourceDescriptionName;
		this.resourceDescriptionUnitOfMeasurement = resourceDescriptionUnitOfMeasurement;
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
