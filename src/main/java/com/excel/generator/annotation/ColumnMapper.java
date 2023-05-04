package com.excel.generator.annotation;

import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;

/**
 * @author 6118454 - [Manish Kumar]
 */
@Retention(RetentionPolicy.RUNTIME)
public @interface ColumnMapper {
    String name();
}