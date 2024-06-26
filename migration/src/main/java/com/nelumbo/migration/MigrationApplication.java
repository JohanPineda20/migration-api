package com.nelumbo.migration;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.cloud.openfeign.EnableFeignClients;

@EnableFeignClients
@SpringBootApplication
public class MigrationApplication {

	public static void main(String[] args) {
		SpringApplication.run(MigrationApplication.class, args);
	}

}
