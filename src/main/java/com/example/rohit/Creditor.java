package com.example.rohit;

public class Creditor {
    private String creditorName;
    private String id;
 
    public Creditor() {
    }
 
    public Creditor(String creditorName, String id) {
        this.setCreditorName(creditorName);
        this.setId(id);
    }

	public String getCreditorName() {
		return creditorName;
	}

	public void setCreditorName(String creditorName) {
		this.creditorName = creditorName;
	}

	public String getId() {
		return id;
	}

	public void setId(String id) {
		this.id = id;
	}
	
}
