package com.pvxy.servicebaseconsume.api;

import org.springframework.cloud.openfeign.FeignClient;
import org.springframework.stereotype.Component;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;

@Component
@FeignClient(value = "service-base-provider", path = "/config")
public interface ConfigApi {

    @GetMapping("/echo/{string}")
    String echo(@PathVariable("string") String string);

}
