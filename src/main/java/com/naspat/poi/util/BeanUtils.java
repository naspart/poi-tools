package com.naspat.poi.util;

import java.lang.reflect.Field;
import java.lang.reflect.Modifier;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

public class BeanUtils {
    public static Field[] getClassFields(Class<?> clazz) {
        List<Field> list = new ArrayList<>();
        Field[] fields;
        do {
            fields = clazz.getDeclaredFields();
            Collections.addAll(list, fields);
            clazz = clazz.getSuperclass();
        } while (clazz != Object.class && clazz != null);

        return list.toArray(fields);
    }

    /**
     * 直接读取对象属性值, 无视private/protected修饰符, 不经过getter函数.
     */
    public static Object getFieldValue(final Object object, final String fieldName) {
        Field field = getDeclaredField(object, fieldName);

        if (field == null) {
            throw new IllegalArgumentException("Could not find field [" + fieldName + "] on target [" + object + "]");
        }

        makeAccessible(field);

        Object result = null;
        try {
            result = field.get(object);
        } catch (IllegalAccessException e) {
        }

        return result;
    }

    /**
     * 循环向上转型, 获取对象的DeclaredField.
     * <p>
     * 如向上转型到Object仍无法找到, 返回null.
     */
    protected static Field getDeclaredField(final Object object,
                                            final String fieldName) {
        for (Class<?> superClass = object.getClass(); superClass != Object.class; superClass = superClass
                .getSuperclass()) {
            try {
                return superClass.getDeclaredField(fieldName);
            } catch (NoSuchFieldException e) {

            }
        }
        return null;
    }

    /**
     * 强行设置Field可访问.
     */
    protected static void makeAccessible(final Field field) {
        if (!Modifier.isPublic(field.getModifiers()) || !Modifier.isPublic(field.getDeclaringClass().getModifiers())) {
            field.setAccessible(true);
        }
    }
}
