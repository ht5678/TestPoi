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
	 * 使用poi读取excel中的数据
	 */
	@Test
	public void testHello(){
		service.testHello();
		
	}
	

}
