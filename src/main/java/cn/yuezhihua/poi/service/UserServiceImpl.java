package cn.yuezhihua.poi.service;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import cn.yuezhihua.poi.domain.Employee;
import cn.yuezhihua.poi.domain.User;
import cn.yuezhihua.poi.util.ExcelAnnotationUtil;
import cn.yuezhihua.poi.util.ExcelTemplate;
import cn.yuezhihua.poi.util.ExcelUtil;

public class UserServiceImpl implements UserService{

	/**
	 * 测试ExtResource注解
	 */
	public void testExcelResource(){
		List<Employee> emps = new ArrayList<Employee>();
		Employee emp1 = new Employee();
		emp1.setId(1);
		emp1.setEname("zhangsan");
		emp1.setEage(11);
		emps.add(emp1);
		
		Employee emp2 = new Employee();
		emp2.setId(2);
		emp2.setEage(12);
		emp2.setEname("wangwu");
		emps.add(emp2);
		
		try{
			OutputStream os = new FileOutputStream(new File("c://poi//emp1.xls"));
	//		ExcelUtil.getInstance().exportObj2Excel("/user.xls", emps, Employee.class, true);
			Map<String, String> datas = new HashMap<String, String>();
			datas.put("title","测试员工信息");
			datas.put("date","2012-06-02 12:33");
			datas.put("dep","昭通师专财务处");
			ExcelUtil.getInstance().exportObj2ExcelByTemplate(datas, "/user.xls", os , emps, Employee.class, true);
//			ExcelUtil.getInstance().exportObj2Excel(os, emps, Employee.class, false);
		}catch(Exception e){
			System.out.println(e.getMessage());
		}
		
	}

	public void testWrite03() {
		ExcelTemplate et = ExcelTemplate.getInstance()
					.readTemplateByClasspath("/default.xls");
		et.createNewRow();
		et.createCell("1111111");
		et.createCell("aaaaaaaaaaaa");
		et.createCell("a1");
		et.createCell("a2a2");
		et.createNewRow();
		et.createCell("222222");
		et.createCell("bbbbb");
		et.createCell("b");
		et.createCell("dbbb");
		et.createNewRow();
		et.createCell(333);
		et.createCell("cccccc");
		et.createCell("a1");
		et.createCell(12333);
		et.createNewRow();
		et.createCell("4444444");
		et.createCell("ddddd");
		et.createCell("a1");
		et.createCell("a2a2");
		et.createNewRow();
		et.createCell("555555");
		et.createCell("eeeeee");
		et.createCell("a1");
		et.createCell(112);
		et.createNewRow();
		et.createCell("555555");
		et.createCell("eeeeee");
		et.createCell("a1");
		et.createCell("a2a2");
		et.createNewRow();
		et.createCell("555555");
		et.createCell("eeeeee");
		et.createCell("a1");
		et.createCell("a2a2");
		et.createNewRow();
		et.createCell("555555");
		et.createCell("eeeeee");
		et.createCell("a1");
		et.createCell("a2a2");
		et.createNewRow();
		et.createCell("555555");
		et.createCell("eeeeee");
		et.createCell("a1");
		et.createCell("a2a2");
		et.createNewRow();
		et.createCell("555555");
		et.createCell("eeeeee");
		et.createCell("a1");
		et.createCell("a2a2");
		et.createNewRow();
		et.createCell("555555");
		et.createCell("eeeeee");
		et.createCell("a1");
		et.createCell("a2a2");
		Map<String,String> datas = new HashMap<String,String>();
		datas.put("title","测试用户信息");
		datas.put("date","2012-06-02 12:33");
		datas.put("dep","昭通师专财务处");
		et.replaceFinalData(datas);
		et.insertSer();
		et.writeToFile("c:/poi/test01.xls");
	}
	
	
	/**
	 * poi---excel的写入测试---对象写入
	 */
	public void testWrite02(){
		
		//初始化user对象
		
		String path = "c:/poi/";
		Workbook wb = new HSSFWorkbook();
		FileOutputStream fos = null;
		try {
			fos = new FileOutputStream(path+"test1.xls");
			//输出文件路径
			System.out.println(path+"test1.xls");
			ExcelAnnotationUtil.Obj2Excel(null, User.class, path+"user.xls");
			wb.write(fos);
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}finally{
			if(fos!=null)
				try {
					fos.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
		}
		
	}
	
	
	/**
	 * poi---excel的写入测试
	 */
	public void testWrite01(){
		//获取项目路径
//		String path = this.getClass().getProtectionDomain().getCodeSource().getLocation().getPath();
		String path = "c:/poi/";
		Workbook wb = new HSSFWorkbook();
		FileOutputStream fos = null;
		try {
			fos = new FileOutputStream(path+"test1.xls");
			//输出文件路径
			System.out.println(path+"test1.xls");
			//创建样式----设置样式
			CellStyle style = wb.createCellStyle();
			style.setAlignment(CellStyle.ALIGN_CENTER);
			style.setAlignment(CellStyle.VERTICAL_CENTER);
			
			//创建sheet
			Sheet sheet = wb.createSheet("test");
			//创建row
			Row row = sheet.createRow(0);
			//创建cell
			Cell cell = row.createCell(0);
			cell.setCellValue("编号");
			cell = row.createCell(1);
			cell.setCellValue("姓名");
			
			wb.write(fos);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}finally{
			if(fos!=null)
				try {
					fos.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
		}
	}
	
	
	/**
	 * poi---excel的读出测试
	 *第一个测试
	 */
	public void testHello(){
		try {
			//获取项目路径
			String path = this.getClass().getProtectionDomain().getCodeSource().getLocation().getPath();
			// Use a file
			Workbook wb = WorkbookFactory.create(new File(path+"user.xls"));
			
			Sheet sheet = wb.getSheetAt(0);
			//for,,,poi中的强循环，和普通的for循环不太一样
			for(Row row : sheet){
				for(Cell cell : row){
					System.out.print(cell.getStringCellValue()+"-----");
				}
				System.out.println();
			}
			
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
}
