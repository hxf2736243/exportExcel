package com.boxer.commom.aop;

import com.boxer.commom.annotation.ExportExcel;
import com.boxer.commom.annotation.TotalRow;
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

@Slf4j
@Aspect
@Component
public class ExportExcelAspect {
    //    @Pointcut("execution(* com..*.*demo.*.*(..)) @annotation(com.boxer.commom.annotation.ExportExcel)")
    @Pointcut("@within(org.springframework.web.bind.annotation.RestController) @annotation(com.boxer.commom.annotation.ExportExcel)")
    public void exportPointcut() {
    }

    @Around("exportPointcut()")
    public Object handleExport(ProceedingJoinPoint joinPoint) throws Throwable {
        log.info("Excel export process started...");
        Object result = joinPoint.proceed();

        if (result instanceof List<?> dataList) {
            MethodSignature signature = (MethodSignature) joinPoint.getSignature();
            Method method = signature.getMethod();
            ExportExcel exportExcel = method.getAnnotation(ExportExcel.class);
            ParameterizedType returnType = (ParameterizedType) method.getGenericReturnType();
            Class<?> clazz = (Class<?>) returnType.getActualTypeArguments()[0];
            HttpServletResponse response = ((ServletRequestAttributes) RequestContextHolder.getRequestAttributes()).getResponse();

            TotalRow totalRow = exportExcel.totalRow();
            // 调用工具类生成Excel
            ExportExcelUtil.export(response, dataList, exportExcel.fileName(), clazz,exportExcel.sameValueMergeColumns(),totalRow.enableTotal(),totalRow.headHeight(),totalRow.totalMergeColumnStart(),totalRow.totalMergeColumnEnd(),totalRow.totalText());
        }
        log.info("Excel export process completed successfully.");
        return result; // 返回null以终止后续逻辑
    }

}
