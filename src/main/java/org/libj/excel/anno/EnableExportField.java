package org.libj.excel.anno;


import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @author eric ling
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface EnableExportField {
    //列宽
    int colWidth() default 100;

    //列名
    String colName();

    //getter
    String getter() default "";

}
