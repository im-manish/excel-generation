package com.excel.generator.annotation;

import org.apache.poi.hssf.util.HSSFColor;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Inherited;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * The Annotation class that handle the information of Excel information put on
 * each variables of the VO Class that require the generation of Excel
 *
 * @author Manish
 */

@Documented
@Target(ElementType.FIELD)
@Inherited
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelInfo {

    /** the header name */
    String headerName() default "No Name";

    /** the sorting index */
    double sortIndex() default 0;

    /** the header color */
    HSSFColor.HSSFColorPredefined headerColor() default HSSFColor.HSSFColorPredefined.YELLOW;

    /** the column width */
    int columnWidth() default 15;

    /** the header color */
    HSSFColor.HSSFColorPredefined headerFontColor() default HSSFColor.HSSFColorPredefined.BLACK;

}
