package com.itcray.test_base.Car;

import java.util.ArrayList;
import java.util.List;

public class CarDTO {

    private List<Car> cars;

    public CarDTO() {
        this.cars = new ArrayList<>();
    }

    public CarDTO(List<Car> cars) {
        this.cars = cars;
    }

    public List<Car> getCars() {
        return cars;
    }

    public void setCars(List<Car> cars) {
        this.cars = cars;
    }

    public void addCar(Car car) {
        this.cars.add(car);
    }
}