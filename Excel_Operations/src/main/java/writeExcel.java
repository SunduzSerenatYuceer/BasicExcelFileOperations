

import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;




public class writeExcel {
	
	static Workbook wb;
	static FileOutputStream fileOut;
	static Sheet sheet;

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		
		createNewFile();
		createSheet();
		createCells();

	}
	
	
	public static void createNewFile() throws Exception {
		
		try {
			
			wb = new XSSFWorkbook();
			
			fileOut = new FileOutputStream(System.getProperty("user.dir") + "/write.xls");
			System.out.println(System.getProperty("user.dir") + "/write.xls");
			//wb.write(fileOut);
			//fileOut.close();
		}catch(Exception e){
						
			
		}
	}
	
	
	public static void createSheet() {
		
		try {
			
			sheet = wb.createSheet("new sheet");
			
			//FileOutputStream fileOut = new FileOutputStream(System.getProperty("user.dir") + "/write.xls");
			//wb.write(fileOut);
			//fileOut.close();

			
			
		}catch(Exception e) {
			
		}
	}
	
	public static void createCells() {
		
		try {
		
			CreationHelper createHelper = wb.getCreationHelper();
			Row row = sheet.createRow((short) 0);
			Cell cell = row.createCell(0);
			cell.setCellValue(1);
			
			row.createCell(1).setCellValue(1.2);
			row.createCell(2).setCellValue(createHelper.createRichTextString("New String"));
			row.createCell(3).setCellValue(true);
			
			wb.write(fileOut);
			fileOut.close();
		}catch(Exception e) {
			
		}
		
		
		
	}

}
