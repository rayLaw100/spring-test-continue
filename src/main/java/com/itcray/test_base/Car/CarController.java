package com.itcray.test_base.Car;

import java.util.List;
import java.util.concurrent.atomic.AtomicLong;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.DeleteMapping;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.PutMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.bind.annotation.RestController;

// @RestController
@Controller
// @RequestMapping("/car")
public class CarController {
 
    private static final String template = "Hello, %s!";
    private final AtomicLong counter = new AtomicLong();

    private final CarRepository carrepo;

    @Autowired
    private CarService carService;

    CarController(CarRepository repo){
        this.carrepo =  repo;
    }

    @ResponseBody
    @GetMapping("/greeting")
    public Car car(@RequestParam(value = "name", defaultValue = "World") String name) {
		return new Car(counter.incrementAndGet(), String.format(template, name));
    }

    @ResponseBody
    @GetMapping("/cars")
    List<Car> all(){
        return carrepo.findAll();
    } 

    @GetMapping("/all")
    public String showAll(Model model){
      model.addAttribute("cars", carrepo.findAll());
      return "allcar.html";
    }

    @ResponseBody
    @PostMapping("/cars")
    Car newCar(@RequestBody Car newCar) {
        return carrepo.save(newCar);
    }
    
    @ResponseBody
    @GetMapping("/cars/{id}")
    Car one(@PathVariable Long id) {

        return  carrepo.findById(id)
        .orElseThrow(()->new CarNotFoundException(id));
    }

    @ResponseBody
    @PutMapping("/cars/{id}")
    Car replaceCar(@RequestBody Car newCar, @PathVariable Long id) {
  
      return carrepo.findById(id)
        .map(car -> {
          car.setName(newCar.getName());
          car.setBrand(newCar.getBrand());
          car.setPrice(newCar.getPrice());
          return carrepo.save(car);
        })

        .orElseGet(() -> {
          newCar.setId(id);
          return carrepo.save(newCar);
        });
    }
    
    @ResponseBody
    @DeleteMapping("/cars/{id}")
    void deleteEmployee(@PathVariable Long id) {
      carrepo.deleteById(id);
    }

}