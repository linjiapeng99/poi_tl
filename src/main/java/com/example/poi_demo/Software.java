package com.example.poi_demo;

public class Software {
    private String swName; // 软件名称
    private String swVersion; // 版本号
    private String swSource; // 生产商/来源
    private String swPurpose; // 用途

    public Software() {
    }

    public Software(String swName, String swVersion, String swSource, String swPurpose) {
        this.swName = swName;
        this.swVersion = swVersion;
        this.swSource = swSource;
        this.swPurpose = swPurpose;
    }

    public String getSwName() {
        return swName;
    }

    public void setSwName(String swName) {
        this.swName = swName;
    }

    public String getSwVersion() {
        return swVersion;
    }

    public void setSwVersion(String swVersion) {
        this.swVersion = swVersion;
    }

    public String getSwSource() {
        return swSource;
    }

    public void setSwSource(String swSource) {
        this.swSource = swSource;
    }

    public String getSwPurpose() {
        return swPurpose;
    }

    public void setSwPurpose(String swPurpose) {
        this.swPurpose = swPurpose;
    }
}
