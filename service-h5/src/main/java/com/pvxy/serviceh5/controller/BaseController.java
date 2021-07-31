package com.pvxy.serviceh5.controller;

import com.pvxy.servicebaseconsume.api.ConfigApi;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import vo.R;

@RestController
@RequestMapping("/base")
public class BaseController {

    @Autowired
    private ConfigApi configApi;

    @GetMapping
    public R test2() {
        return R.success(configApi.echo("nnnn"));
    }

}
