package com.adroitwolf.util;

import lombok.extern.slf4j.Slf4j;
import org.springframework.util.StringUtils;

import java.lang.reflect.Method;


/**
 * <pre>ReflectUtil</pre>
 *
 * @author adroitwolf 2021年03月16日 15:23
 */
@Slf4j
public class ReflectUtil {

    private static final String GETTER_PREFIX = "get";


    /**
     * 通过Get方法
     */
    public static Method getMethod(Class<?> sourceClass, String propertyName)  {


        String getterMethodName = GETTER_PREFIX + StringUtils.capitalize(propertyName);
        Method method= null;
        try {
            method = sourceClass.getMethod(getterMethodName);
        } catch (NoSuchMethodException e) {
            log.error("没有该方法！");
        }

        return method;
    }



    public static Object invokeMethod(final Object obj, final Method method) {
        if (obj == null || method == null) {
            return null;
        }
        try {
            return  method.invoke(obj);
        } catch (Exception e) {
            String msg = "method: " + method + ", obj: " + obj +  "";
            log.error(msg+"找不到");

        }
        return null;
    }


}
