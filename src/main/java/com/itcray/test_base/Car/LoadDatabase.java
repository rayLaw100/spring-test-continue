package com.itcray.test_base.Car;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.boot.CommandLineRunner;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;


@Configuration
class LoadDatabase {

  private static final Logger log = LoggerFactory.getLogger(LoadDatabase.class);

  @Bean
  CommandLineRunner initDatabase(CarRepository repository) {

    return args -> {
      log.info("Preloading " + repository.save(new Car("Bilbo Baggins4", "Fatal",145)));
      log.info("Preloading " + repository.save(new Car("Frodo RUNNER", "BMW",133)));
      log.info("Preloading 3" + repository.save(new Car("Try3", "BMW",134)));
    };
  }
}