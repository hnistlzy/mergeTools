package com.pactera.pojo;

public class User  {
   private String username;
    private String bugName;
    private String date;
    private  String severity;
    private String testWay;

    public String getUsername() {
        return username;
    }

    public void setUsername(String username) {
        this.username = username;
    }

    public String getBugName() {
        return bugName;
    }

    public void setBugName(String bugName) {
        this.bugName = bugName;
    }

    public String getDate() {
        return date;
    }

    public void setDate(String date) {
        this.date = date;
    }

    public String getSeverity() {
        return severity;
    }

    public void setSeverity(String severity) {
        this.severity = severity;
    }

    public String getTestWay() {
        return testWay;
    }

    public void setTestWay(String testWay) {
        this.testWay = testWay;
    }
}
