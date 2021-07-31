package com.pvxy.servicebaseprovider;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.cloud.client.discovery.EnableDiscoveryClient;

@EnableDiscoveryClient
@SpringBootApplication
public class ServiceBaseProviderApplication {

    public static void main(String[] args) {
        SpringApplication.run(ServiceBaseProviderApplication.class, args);
    }

}
