package cn.yuezhihua.poi;

import org.junit.Test;

import cn.yuezhihua.poi.service.UserServiceImpl;


/**
 * poi-------excel的测试
 * @author lucas
 *
 */
public class TestExcel {
	
	private UserServiceImpl service = new UserServiceImpl();
	
	/**
	 * 测试ExcelResource注解
	 */
	@Test
	public void testExcelResource(){
		service.testExcelResource();
		
	}
	
	
	/**
	 * poi---excel的自定义注解
	 */	
	@Test
	public void testWrite03(){
		service.testWrite03();
	}
	

	/**
	 * poi---excel的自定义注解
	 */	
	@Test
	public void testWrite02(){
		service.testWrite02();
	}

	/**
	 * poi---excel的写入测试
	 */
	@Test
	public void testWrite01(){
		service.testWrite01();
	}
	
	
	/**
	 * 使用poi读取excel中的数据
	 */
	@Test
	public void testHello(){
		service.testHello();
		
	}
	

}
