package org.labj.excel.anno;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @author eric ling
 */
// 导入索引 从0开始
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ImportIndex {
    //索引
    int index();

    //设置set方法
    String setter() default "";

}
