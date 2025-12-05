package com.example.poi_demo;

import java.util.List;

public class ServEnv {
    private Integer envId;//编号
    private String hwName; // 硬件名称/型号
    private String hwConfig; // 硬件配置
    private String swName;//软件名称
    private String swVersion;//软件版本
    private String swSource;//软件来源
    private String swPurpose;//软件用途
    private List<Software> swList; // 该硬件对应的多个软件环境

    public String getHwConfig() {
        return hwConfig;
    }

    public void setHwConfig(String hwConfig) {
        this.hwConfig = hwConfig;
    }

    public List<Software> getSwList() {
        return swList;
    }

    public void setSwList(List<Software> swList) {
        this.swList = swList;
    }

    public ServEnv() {
    }

    public Integer getEnvId() {
        return envId;
    }

    public void setId(Integer envId) {
        this.envId = envId;
    }

    public ServEnv(Integer envId, String hwName, String hwConfig, String swName, String swVersion, String swSource, String swPurpose) {
        this.envId = envId;
        this.hwName = hwName;
        this.hwConfig = hwConfig;
        this.swName = swName;
        this.swVersion = swVersion;
        this.swSource = swSource;
        this.swPurpose = swPurpose;
    }

    public ServEnv(Integer envId, String hwName, String hwConfig, List<Software> swList) {
        this.envId = envId;
        this.hwName = hwName;
        this.hwConfig = hwConfig;
        this.swList = swList;
    }

    public String getHwName() {
        return hwName;
    }

    public void setHwName(String hwName) {
        this.hwName = hwName;
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
