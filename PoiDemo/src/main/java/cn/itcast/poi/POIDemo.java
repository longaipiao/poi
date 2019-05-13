package cn.itcast.poi;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

public class POIDemo {

	// 写excel
	@Test
	public void writeExcel() throws IOException {
		//1.创建一个输出流
		FileOutputStream fos=new FileOutputStream("d:/poi.xlsx");		
		//2.创建一个Workbook来指定对文件操作
		XSSFWorkbook book=new XSSFWorkbook();
		//2.1创建 一个Sheet
		XSSFSheet sheet = book.createSheet("学生信息");
		
		//2.2创建指定sheet的row
		XSSFRow row = sheet.createRow(0);
		row.createCell(0).setCellValue("姓名");
		row.createCell(1).setCellValue("年龄");
		row.createCell(2).setCellValue("性别");
		
		for(int i=1;i<=3;i++){
			XSSFRow rowi = sheet.createRow(i);
			rowi.createCell(0).setCellValue("姓名"+i);
			rowi.createCell(1).setCellValue(i+20);
			rowi.createCell(2).setCellValue("性别"+i);
		}		
		//3.写文件
		book.write(fos);		
		fos.close();
	}

	// 读excel
	@Test
	public void readFExcel() throws IOException {
		//1.创建一个输入流
		FileInputStream fis=new FileInputStream("d:/poi.xlsx");		
		//2.创建一个book
		XSSFWorkbook book=new XSSFWorkbook(fis);		
		//3.操作
		XSSFSheet sheet = book.getSheetAt(0);		
		// XSSFRow row = sheet.getRow(0);
		// XSSFCell cell = row.getCell(0);
		// System.out.println(cell.getStringCellValue());
		
		//需求:将指定的excel文件中的内容打印到控制台上.
		for(int i=0;i<=sheet.getLastRowNum();i++){
			XSSFRow row = sheet.getRow(i);
			
			for(int j=0;j<row.getLastCellNum();j++){
				
				XSSFCell cell = row.getCell(j);
				
				switch(cell.getCellType()){
					case Cell.CELL_TYPE_NUMERIC:
						System.out.print(cell.getNumericCellValue()+"\t");
						break;
					case Cell.CELL_TYPE_STRING:
						System.out.print(cell.getStringCellValue()+"\t");
						break;
				}
				
			}
			System.out.println();		
		}		
		//4.关闭
		fis.close();
	}
}
