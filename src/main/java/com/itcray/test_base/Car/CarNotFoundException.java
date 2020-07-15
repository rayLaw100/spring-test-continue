package com.itcray.test_base.Car;


public class CarNotFoundException extends RuntimeException {

    /**
     *
     */
    private static final long serialVersionUID = -5747805141125199253L;

    CarNotFoundException(Long id) {
      super("Could not find any car with id: " + id);
    }
}