import java.io.FileOutputStream;
import java.util.Random;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelGenerator {

	public static void main(String[] args) {
		
		String[] names = {"Ivo", "Ema", "Asen", "Lora", "Dany", "Vasko", "Gergana", "Georgi", "Mitko", "Ana"};
		
		XSSFWorkbook wb = new XSSFWorkbook();
//		Workbook wb = new HSSFWorkbook();
		Sheet sheet1 = wb.createSheet();
		
		Row row1 = sheet1.createRow(0);
		
		CellStyle headerStyle = wb.createCellStyle();
		XSSFFont headerFont = (XSSFFont) wb.createFont();
	    headerFont.setBold(true);
	    
		headerStyle.setFillForegroundColor(IndexedColors.AQUA.getIndex());
		headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	    headerStyle.setFont(headerFont);
	    
		row1.setRowStyle(headerStyle);
		
		row1.createCell(0).setCellValue("Name");
		row1.getCell(0).setCellStyle(headerStyle);
		row1.createCell(1).setCellValue("Age");
		row1.getCell(1).setCellStyle(headerStyle);
		row1.createCell(2).setCellValue("Score");
		row1.getCell(2).setCellStyle(headerStyle);
		
		CellStyle oddRowStyle = wb.createCellStyle();
		XSSFFont oddRowFont = wb.createFont();
		oddRowFont.setColor(IndexedColors.GREEN.getIndex());
		oddRowStyle.setFont(oddRowFont);
		
		Random rand = new Random();
		
		for (int i = 0; i < 100000; i++){
			Row row = sheet1.createRow(i + 1);
			row.createCell(0).setCellValue(names[rand.nextInt(10)]);
			row.createCell(1).setCellValue(rand.nextInt(81) + 20);
			row.createCell(2).setCellValue(rand.nextInt(101));
			
			if(i % 2 == 1){
				row.getCell(0).setCellStyle(oddRowStyle);
				row.getCell(1).setCellStyle(oddRowStyle);
				row.getCell(2).setCellStyle(oddRowStyle);
			}
		}
		
		row1.createCell(4).setCellValue("Average Score");
		row1.getCell(4).setCellStyle(headerStyle);
		sheet1.getRow(1).createCell(4).setCellFormula("AVERAGE(C2:C101)");
	
		try {
			FileOutputStream output = new FileOutputStream("scores.xlsx");
			wb.write(output);
			output.close();
			
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
