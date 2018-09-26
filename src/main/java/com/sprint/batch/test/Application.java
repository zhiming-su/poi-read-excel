package com.sprint.batch.test;

import java.io.FileNotFoundException;
import java.io.IOException;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class Application {

	public static void main(String[] args) throws FileNotFoundException, IOException {
		SpringApplication.run(Application.class, args);
		readExcel.read();
	}
}
