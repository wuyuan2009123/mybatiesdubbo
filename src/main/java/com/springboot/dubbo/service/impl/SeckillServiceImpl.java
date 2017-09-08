package com.springboot.dubbo.service.impl;

/**
 * Created by wuy on 2017/9/4.
 */


import com.springboot.dubbo.entity.Seckill;
import com.springboot.dubbo.mapper.SeckillDao;
import com.springboot.dubbo.service.SeckillService;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.List;



@Service
public class SeckillServiceImpl implements SeckillService {

    private Logger logger = LoggerFactory.getLogger(this.getClass());

    // 注入dao依赖
    @Autowired
    private SeckillDao seckillDao;


    public List<Seckill> getSeckillList() {
        return seckillDao.queryAll(0, 4);
    }

    public Seckill getById(long seckillId) {
        return seckillDao.queryById(seckillId);
    }





}
