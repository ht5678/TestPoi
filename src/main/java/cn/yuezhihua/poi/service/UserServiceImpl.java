package cn.yuezhihua.poi.service;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import cn.yuezhihua.poi.domain.User;
import cn.yuezhihua.poi.util.ExcelUtil;

public class UserServiceImpl implements UserService{

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
			ExcelUtil.Obj2Excel(null, User.class, path+"user.xls");
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
