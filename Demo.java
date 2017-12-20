package cn.itcast.json.test;

import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

public class Demo {
	@SuppressWarnings("resource")
	@Test
	public void run() throws Throwable{
		//创建工作薄
		Workbook wb = new HSSFWorkbook();
		//创建工作表
		Sheet createSheet = wb.createSheet();
		//创建第2行
		Row row = createSheet.createRow(1);
		//创建单元格第2列
		Cell cell = row.createCell(1);
		//在单元格里写入数据
		cell.setCellValue("世纪东方静安寺");
		//创建输出流
		OutputStream os = new FileOutputStream("I://abc.xls");
		wb.write(os);
		//关闭流
		wb.close();
		os.close();
	}
}
