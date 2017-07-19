package com.kjubo.excel;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelColumn {

	/**
	 * excel列顺序值
	 *
	 * @return
	 */
	int col();

	/**
	 * coding 工具类
	 *
	 * @return
	 */
	Class<? extends ICodeable> coding() default ICodeable.None.class;

	/**
	 * 名称
	 *
	 * @return
	 */
	String name();

}