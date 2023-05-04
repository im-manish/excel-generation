package com.excel.generator.util;

import com.excel.generator.annotation.ColumnMapper;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.*;

/**
 * @author 6118454 - [Manish Kumar]
 */
public class RowMapperUtils<T> {

    private static final Logger log = LogManager.getLogger(RowMapperUtils.class);

    private Class<T> clazz;
    private final Map<String, Field> fields = new HashMap<>();
    Map<String, String> errors = new HashMap<>();

    public RowMapperUtils(Class clazz) {
        this.clazz = clazz;
        List<Field> fieldList = Arrays.asList(clazz.getDeclaredFields());
        for (Field field : fieldList) {
            ColumnMapper columnMapper = field.getAnnotation(ColumnMapper.class);
            if (columnMapper != null) {
                field.setAccessible(true);
                fields.put(columnMapper.name(), field);
            }
        }
    }

    public List<T> mapDataObjectToDTO(ResultSet resultSet) throws SQLException, NoSuchMethodException, IllegalAccessException, InvocationTargetException, InstantiationException {
        List<T> list = new ArrayList<>();
        while (resultSet.next()) {
            T dto = (T) clazz.getConstructor().newInstance();
            for (Map.Entry<String, Field> entity : fields.entrySet()) {
                String columnName = entity.getKey();
                Field field = entity.getValue();
                try {
                    Object object = resultSet.getObject(columnName, field.getType());
                    if (Objects.isNull(object)) {
                        continue;  // Don't set DBNULL
                    }
                    field.set(dto, convertInstanceOfObject(object));
                } catch (Exception e) {
                    log.error("Error while setting the value for columnName : {}", columnName,e);
                }
            }
            list.add(dto);
        }
        return list;
    }

    private T convertInstanceOfObject(Object o) {
        try {
            return (T) o;
        } catch (ClassCastException e) {
            log.error("Error while converting the Instance of Object", e);
            return null;
        }
    }
}