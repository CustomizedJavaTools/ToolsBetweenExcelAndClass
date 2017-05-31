package com.hsbc.gbm.ptutilities.common.util.excelTools;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;
/**
 * ExcelColumn.java
 * Purpose:	configuration of mapping field in class with excel column.
 *
 * @author	Herman H WEN
 * @version 1.0 2017/05/16
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelColumn {
	public String column() default "";
	public String header() default "";
	public int maxLength() default -1;
	public int minLength() default 0;
	public boolean required() default false;
}
