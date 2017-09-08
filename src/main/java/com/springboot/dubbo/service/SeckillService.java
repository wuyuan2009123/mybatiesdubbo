package com.springboot.dubbo.service;


import com.springboot.dubbo.entity.Seckill;

import java.util.List;


/**
 * Created by wuy on 2017/9/4.
 */
public interface SeckillService {

    /**
     * 查询所有秒杀记录
     *
     * @return
     */
    List<Seckill> getSeckillList();

    /**
     * 查询单个秒杀记录
     *
     * @param seckillId
     * @return
     */
    Seckill getById(long seckillId);


}
