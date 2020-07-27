package it.raffles.cimc.excel.annotation;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

import com.alibaba.excel.support.ExcelTypeEnum;
import com.alibaba.excel.write.handler.WriteHandler;

@Documented
@Target(ElementType.METHOD)
@Retention(RetentionPolicy.RUNTIME)
public @interface Excel {

	String contentType() default "application/vnd.ms-excel";
	
	String charset() default "utf-8";

	ExcelTypeEnum suffix() default ExcelTypeEnum.XLSX;

	String sheetName() default "sheet";
	
	String fileName() default "export";
	
	String templateFileName() default "";
	
	String templateFolderName() default "excel";
	
	String[] columnNames() default {};
	
	String password() default "";
	
	boolean isLongestMatchColumnWidthStyleStrategy() default true ;
	
	Class<?>[] entity() default {};
	
	Class<? extends WriteHandler>[] writeHandlers() default {};
}
