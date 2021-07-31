package com.pvxy.servicebaseprovider.controller;

import org.springframework.cloud.context.config.annotation.RefreshScope;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("/config")
@RefreshScope
public class ConfigController {

    @GetMapping("/echo/{string}")
    public String echo(@PathVariable("string") String string) {
        return "Hello Nacos Discovery " + string;
    }
}
