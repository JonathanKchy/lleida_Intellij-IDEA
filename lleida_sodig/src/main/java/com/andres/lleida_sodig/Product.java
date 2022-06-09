package com.andres.lleida_sodig;

public class Product {
    private String name;
    private String age;
    private String animal;

    public Product(String name, String age, String animal) {
        this.name = name;
        this.age = age;
        this.animal = animal;
    }

    public String getName() {
        return name;
    }

    public String getAge() {
        return age;
    }

    public String getAnimal() {
        return animal;
    }
}
