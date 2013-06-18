package cn.yuezhihua.poi.util;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;


/**
 * @ExcelUtil注解的使用关系类
 * @author yuezhihua
 *
 */
public class ExcelUtil {

	/**
	 * 将一个用户集合输出到制定的path路径，并且生成excel文件
	 * @param <T>
	 * @param T
	 * @param clazz
	 * @param path
	 */
	public static <T> void Obj2Excel(List<T> T , Class<T> clazz , String path){
		Workbook wb = new HSSFWorkbook();
		FileOutputStream fos = null;
		try {
			fos = new FileOutputStream(new File(path));
			cn.yuezhihua.poi.annotation.ExcelUtil anno = (cn.yuezhihua.poi.annotation.ExcelUtil) clazz
					.getAnnotation(cn.yuezhihua.poi.annotation.ExcelUtil.class);
			Field[] columns = clazz.getDeclaredFields();
			Sheet sheet = wb.createSheet("excelutil注解01");
			CellStyle style = wb.createCellStyle();
			//写入标题
			Row row = sheet.createRow(0);
			
			Cell cell = row.createCell(0);
			//合并单元格
			sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, columns.length-1));
			if(anno.title()!=null){
				cell.setCellValue(anno.title());
			}
			//写入字段名称
			row = sheet.createRow(1);
			for(int i = 0 ; i < columns.length ;i++){
				cell = row.createCell(i);
				anno = (cn.yuezhihua.poi.annotation.ExcelUtil)columns[i].getAnnotation(cn.yuezhihua.poi.annotation.ExcelUtil.class);
				cell.setCellValue(anno.cname());
			}
//			//写入数据
//			for(int i = 0 ; i < T.size() ; i++){
//				
//			}
			wb.write(fos);
		} catch (Exception e) {
			System.out.println("==="+e.getMessage());
		} finally {
			if (fos != null)
				try {
					fos.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
		}

	}
	
}
