package club.xsir.poi;

import java.lang.reflect.Field;
import java.util.*;
import java.util.stream.Collectors;

public class FieldOrderUtil {
    
    public static List<Field> getFieldsInOrder(Class<?> clazz) {
        return Arrays.stream(clazz.getDeclaredFields())
                .sorted(Comparator.comparingInt(field -> {
                    FieldInfo order = field.getAnnotation(FieldInfo.class);
                    return order != null ? order.order() : Integer.MAX_VALUE;
                }))
                .collect(Collectors.toList());
    }

    public static Map<Integer,Field> getFieldsInOrderMap(Class<?> clazz) {
        return Arrays.stream(clazz.getDeclaredFields()).map(field -> {
            FieldInfo order = field.getAnnotation(FieldInfo.class);
            if (null == order){
                return null;
            }
            Map<Integer, Field> result = new HashMap<>();
            result.put(order.order(),field);
            return result;
        }).filter(Objects::nonNull).reduce((a,b)-> {
            a.putAll(b);
            return a;
        }).orElse(new HashMap<>());
    }
    
    public static void printFieldsInOrder(Class<?> clazz) throws InstantiationException, IllegalAccessException {
        List<Field> orderedFields = getFieldsInOrder(clazz);
        Object o = clazz.newInstance();
        System.out.println("字段顺序:");
        for (Field field : orderedFields) {
            FieldInfo order = field.getAnnotation(FieldInfo.class);
            if (null != order){
                field.setAccessible(true);
                field.set(o,String.format("%s-%s",field.getName(),order.order()));
                field.setAccessible(false);
                System.out.printf("%s (order=%d)%n",
                        field.getName(),
                        order.order());
            }
        }
        System.out.println(o);
    }

    public static void main(String[] args) throws InstantiationException, IllegalAccessException {
        printFieldsInOrder(DataStruct.class);
    }
}