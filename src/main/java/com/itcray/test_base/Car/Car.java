package com.itcray.test_base.Car;

import java.util.Objects;

import javax.persistence.Entity;
import javax.persistence.GeneratedValue;
import javax.persistence.Id;
import javax.persistence.Table;

import lombok.Getter;
import lombok.Setter;

@Entity
@Getter @Setter
@Table
public class Car {
    

    private @Id @GeneratedValue Long id;
    private String name;
    private String brand;
    private int price;

    public Car(){

    }

    public Car(Long id, String name){
        this.id =id;
        this.name =name;
    }

    public Car(String name, String brand,int price){
        this.brand =brand;
        this.name =name;
        this.price =price;
    }

    @Override
    public boolean equals(Object o){

        if(this==o){
            return true;

        }
        if (!(o instanceof Car))
        return false;
        Car car = (Car) o ;
        return Objects.equals(this.id, car.id)&& 
        Objects.equals(this.name, car.name)&&
        Objects.equals(this.brand, car.brand)&&
        Objects.equals(this.price, car.price);

    }

    @Override
    public int hashCode() {
      return Objects.hash(this.id, this.name, this.brand, this.price);
    }
  
    @Override
    public String toString() {
      return "CAR{" + "id=" + this.id + ", name='" + this.name + '\'' + ", role='" + this.brand + '\'' + ", price='" + this.price + '\'' + '}';
    }

}