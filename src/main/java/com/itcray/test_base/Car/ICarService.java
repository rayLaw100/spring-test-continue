package com.itcray.test_base.Car;

import org.springframework.stereotype.Service;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.function.Function;
import java.util.stream.Collectors;

@Service
public class ICarService implements CarService {

    static Map<Long, Car> CarsDB = new HashMap<>();

    @Override
    public List<Car> findAll() {
        return new ArrayList<>(CarsDB.values());
    }

    // @Override
    // public void saveAll(List<Book> books) {
    //     long nextId = getNextId();
    //     for (Book book : books) {
    //         if (book.getId() == 0) {
    //             book.setId(nextId++);
    //         }
    //     }

    //     Map<Long, Book> bookMap = books.stream()
    //         .collect(Collectors.toMap(Book::getId, Function.identity()));

    //     booksDB.putAll(bookMap);
    // }

    private Long getNextId() {
        return CarsDB.keySet()
            .stream()
            .mapToLong(value -> value)
            .max()
            .orElse(0) + 1;
    }
}