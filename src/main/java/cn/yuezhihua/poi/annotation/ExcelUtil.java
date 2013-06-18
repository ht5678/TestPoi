package cn.yuezhihua.poi.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;


/**
 * 自定义注解
 * @author yuezhihua
 *
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(value={ElementType.METHOD,ElementType.TYPE,ElementType.FIELD})
public @interface ExcelUtil {

	/**
	 * 标题
	 * @return
	 */
	public abstract String title() default "";
	
	/**
	 * 字段名称
	 * @return
	 */
	public abstract String cname() default "";
	
	
	
}
