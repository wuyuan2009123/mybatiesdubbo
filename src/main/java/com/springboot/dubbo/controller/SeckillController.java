package com.springboot.dubbo.controller;


import com.springboot.dubbo.entity.Seckill;
import com.springboot.dubbo.service.SeckillService;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;

import java.util.List;
import java.util.concurrent.Callable;

/**
 * Created by wuy on 2017/9/4.
 */
@RestController
@RequestMapping("/seckill") // url:/模块/资源/{id}/细分 /seckill/list
public class SeckillController {

    private Logger logger = LoggerFactory.getLogger(this.getClass());

    @Autowired
    private SeckillService seckillService;

    @RequestMapping(value = "/list", method = RequestMethod.GET)
    public  List<Seckill> list() {
        // 获取列表页
        List<Seckill> list = seckillService.getSeckillList();
        return list;
    }


    @RequestMapping(value = "/{seckillId}/detail", method = RequestMethod.GET)
    public Seckill detail(@PathVariable("seckillId") Long seckillId, Model model) {
        if (seckillId == null) {
            return new Seckill();
        }
        Seckill seckill = seckillService.getById(seckillId);
        return seckill;
    }


    @RequestMapping(value = "/{seckillId}/ayscdetail", method = RequestMethod.GET)
    public Callable<Seckill> ayscdetail(@PathVariable("seckillId") Long seckillId, Model model) {
        Callable<Seckill> result=null;
        if (seckillId == null) {
            result=new Callable<Seckill>() {
                public Seckill call() throws Exception {
                    return new Seckill();
                }
            };
        }
        final Seckill seckill = seckillService.getById(seckillId);
        return new Callable<Seckill>() {
            public Seckill call() throws Exception {
                return seckill;
            }
        };
    }





}