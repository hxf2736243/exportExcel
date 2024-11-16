package com.boxer.commom.aop;

import com.boxer.commom.annotation.ExportExcel;
import com.boxer.commom.utils.ExportExcelUtil;
import jakarta.servlet.http.HttpServletResponse;
import lombok.extern.slf4j.Slf4j;
import org.aspectj.lang.ProceedingJoinPoint;
import org.aspectj.lang.annotation.Around;
import org.aspectj.lang.annotation.Aspect;
import org.aspectj.lang.annotation.Pointcut;
import org.aspectj.lang.reflect.MethodSignature;
import org.springframework.stereotype.Component;
import org.springframework.web.context.request.RequestContextHolder;
import org.springframework.web.context.request.ServletRequestAttributes;

import java.lang.reflect.Method;
import java.lang.reflect.ParameterizedType;
import java.util.List;
import java.util.Objects;

@Slf4j
@Aspect
@Component
public class ExportExcelAspect {
    @Pointcut("@annotation(com.boxer.commom.annotation.ExportExcel)")
    public void aroundPointcut(){}

    @Around(value="aroundPointcut()")
    public Object handleExport(ProceedingJoinPoint joinPoint)throws Throwable
    {
        log.info("handleExport------");
        // 执行目标方法，获取数据
        Object result = joinPoint.proceed();

        MethodSignature signature = (MethodSignature) joinPoint.getSignature();
        Method method = signature.getMethod();
        ParameterizedType returnType = (ParameterizedType) method.getGenericReturnType();
        Class<?> clazz = (Class<?>) returnType.getActualTypeArguments()[0];
        ExportExcel exportExcel=method.getAnnotation(ExportExcel.class);
        HttpServletResponse response=((ServletRequestAttributes) Objects.requireNonNull(RequestContextHolder.getRequestAttributes())).getResponse();

        if (result instanceof List<?>) {
            List<?> dataList = (List<?>) result;
            String fileName = exportExcel.fileName();

            ExportExcelUtil.export(response, dataList, fileName,clazz);
            return null;
        }else{
            return result;
        }
    }
}
