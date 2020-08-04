package Demo;

import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Utilsjava {
	static String path;
	static XSSFWorkbook workbook;
	static XSSFSheet sheet;

	public Utilsjava(String excelpath, String Sheetname) {// ..............Constructor
		try {
			workbook = new XSSFWorkbook(excelpath);
			sheet = workbook.getSheet(Sheetname);
		} catch (Exception e) {
			System.out.println(e.getMessage());
			e.printStackTrace();
		}
	}

	public static void main(String[] args) {
		// System.out.println("hii");
		// getRowCount();
		// getCellCount();
		getCelldataString(0, 0);
		getCelldataNumber(1, 1);
	}

	public static void getRowCount() {
		try {
			int rowc = sheet.getPhysicalNumberOfRows();
			System.out.println("rowcount: " + rowc);
		} catch (Exception e) {
			System.out.println(e.getMessage());
			e.printStackTrace();
		}
	}

	public static void getCellCount() {
		try {
			int cellcount = sheet.getRow(0).getPhysicalNumberOfCells();
			System.out.println("cellcount: " + cellcount);
		} catch (Exception e) {
			System.out.println(e.getMessage());
			e.printStackTrace();
		}
	}

	public static void getCelldataNumber(int rowCount, int colCount) {
		try {
			double num = sheet.getRow(rowCount).getCell(colCount).getNumericCellValue();
			System.out.println("numData :" + num);
		} catch (Exception e) {
			System.out.println(e.getMessage());
			e.printStackTrace();
		}
	}

	public static void getCelldataString(int rowC, int colC) {
		try {
			String stringdata = sheet.getRow(rowC).getCell(colC).getStringCellValue();
			System.out.println("String data:" + stringdata);
		} catch (Exception e) {
			System.out.println(e.getMessage());
			e.printStackTrace();
		}
	}

}
