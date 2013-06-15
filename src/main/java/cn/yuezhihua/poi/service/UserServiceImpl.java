package cn.yuezhihua.poi.service;

import java.io.File;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class UserServiceImpl implements UserService{

	
	/**
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
