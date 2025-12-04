package com.example.poi_demo;

public class Person {
    String name;
    int age;
    String gender;
    String company;
    String remark;

    public Person() {
    }

    public Person(String name, int age, String gender, String company, String remark) {
        this.name = name;
        this.age = age;
        this.gender = gender;
        this.company = company;
        this.remark = remark;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public int getAge() {
        return age;
    }

    public void setAge(int age) {
        this.age = age;
    }

    public String getGender() {
        return gender;
    }

    public void setGender(String gender) {
        this.gender = gender;
    }

    public String getCompany() {
        return company;
    }

    public void setCompany(String company) {
        this.company = company;
    }

    public String getRemark() {
        return remark;
    }

    public void setRemark(String remark) {
        this.remark = remark;
    }
}
