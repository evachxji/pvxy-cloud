package com.pvxy.serviceh5;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.cloud.client.discovery.EnableDiscoveryClient;
import org.springframework.cloud.openfeign.EnableFeignClients;
import org.springframework.context.annotation.ComponentScan;

@SpringBootApplication
@EnableDiscoveryClient
@EnableFeignClients(basePackages = {"com.pvxy.servicebaseconsume"})
public class ServiceH5Application {

	public static void main(String[] args) {
		SpringApplication.run(ServiceH5Application.class, args);
	}

}
