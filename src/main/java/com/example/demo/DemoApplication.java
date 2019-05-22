package com.example.demo;

import com.example.demo.controller.HomeController;
import com.example.demo.repository.ClientRepository;
import com.example.demo.service.ClientService;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class DemoApplication {

    public static void main(String[] args) {
        SpringApplication.run(DemoApplication.class, args);
    }
}
