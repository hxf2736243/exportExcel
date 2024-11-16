package com.boxer.commom.aop;

import lombok.extern.slf4j.Slf4j;
import org.aspectj.lang.ProceedingJoinPoint;
import org.aspectj.lang.annotation.Around;
import org.aspectj.lang.annotation.Aspect;
import org.springframework.stereotype.Component;



@Slf4j
@Aspect
@Component
public class LoggingAspect {
    // 定义一个切点：拦截所有 Controller 层的 public 方法
    @Around("@annotation(com.boxer.commom.annotation.LogExecutionTime)")
    public Object logExecutionTime(ProceedingJoinPoint joinPoint) throws Throwable {
        long start = System.currentTimeMillis();
        Object proceed=null;
        try {
            proceed= joinPoint.proceed();
        } catch (Throwable e) {
            e.printStackTrace();
        }
        long executionTime = System.currentTimeMillis() - start;
       log.info(joinPoint.getSignature() + " executed in " + executionTime + "ms");
        return proceed;
    }
}
