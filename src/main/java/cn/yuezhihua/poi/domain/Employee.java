package cn.yuezhihua.poi.domain;

import java.io.Serializable;

import cn.yuezhihua.poi.util.ExcelResources;

/**
 * 测试ExcelResource注解
 * @author byht
 *
 */
public class Employee implements Serializable{
	

	
	private Integer id;
	
	
	private String ename;
	
	
	private Integer eage;
	
	
	
	@ExcelResources(title="用户ID",order=1)
	public Integer getId() {
		return id;
	}

	public void setId(Integer id) {
		this.id = id;
	}
	@ExcelResources(title="员工名称",order=3)
	public String getEname() {
		return ename;
	}

	public void setEname(String ename) {
		this.ename = ename;
	}
	@ExcelResources(title="员工年龄",order=2)
	public Integer getEage() {
		return eage;
	}

	public void setEage(Integer eage) {
		this.eage = eage;
	}


	
}
