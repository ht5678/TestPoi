package cn.yuezhihua.poi.domain;

import java.io.Serializable;

import cn.yuezhihua.poi.annotation.ExcelUtil;

/**
 * 
 * @author yuezhihua
 *
 */
@ExcelUtil(title="用户信息列表")
public class User implements Serializable{

	
	@ExcelUtil(cname="编号")
	public Integer id;
	
	@ExcelUtil(cname="用户名")
	private String uname;
	
	@ExcelUtil(cname="密码")
	private String upass;
	
	
	public Integer getId() {
		return id;
	}

	public void setId(Integer id) {
		this.id = id;
	}

	public String getUname() {
		return uname;
	}

	public void setUname(String uname) {
		this.uname = uname;
	}

	public String getUpass() {
		return upass;
	}

	public void setUpass(String upass) {
		this.upass = upass;
	}

	
}
