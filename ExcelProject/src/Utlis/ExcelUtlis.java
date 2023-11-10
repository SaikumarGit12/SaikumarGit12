package Utlis;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtlis {

	public static FileInputStream fi;
	public static Workbook wb;
	public static Sheet ws;
	public static Row row;
	public static Cell cell;
	public static FileOutputStream fo;
	public static CellStyle style;

	public static int getRowCount(String Xlfile, String Xlsheet) throws IOException {
		fi = new FileInputStream(Xlfile);
		wb = new XSSFWorkbook(fi);
		ws = wb.getSheet(Xlsheet);
		int rowcount = ws.getLastRowNum();
		wb.close();
		return rowcount;
	}

	public static short getColumnCount(String Xlfile, String Xlsheet, int rowcount) throws IOException {

		fi = new FileInputStream(Xlfile);
		wb = new XSSFWorkbook(fi);
		ws = wb.getSheet(Xlsheet);
		row = ws.getRow(rowcount);
		short colcount = row.getLastCellNum();
		wb.close();
		return colcount;

	}

	public static String getStringData(String Xlfile, String Xlsheet, int rowcount, int colcount) throws IOException {
		fi = new FileInputStream(Xlfile);
		wb = new XSSFWorkbook(fi);
		ws = wb.getSheet(Xlsheet);
		row = ws.getRow(rowcount);
		cell = row.getCell(colcount);
		String data;
		try {
			data = cell.getStringCellValue();
		} catch (Exception e) {
			data = "  ";
		}
		return data;

	}

	public static double getNumericData(String Xlfile, String Xlsheet, int rowcount, int colcount) throws IOException {

		fi = new FileInputStream(Xlfile);
		wb = new XSSFWorkbook(fi);
		ws = wb.getSheet(Xlsheet);
		row = ws.getRow(rowcount);
		cell = row.getCell(colcount);
		Double data;

		try {
			data = cell.getNumericCellValue();
		} catch (Exception e) {
			data = 0.0;
		}
		wb.close();
		return data;

	}

	public static boolean getBooleanData(String Xlfile, String Xlsheet, int rowcount, int colcount) throws IOException {

		fi = new FileInputStream(Xlfile);
		wb = new XSSFWorkbook(fi);
		ws = wb.getSheet(Xlsheet);
		row = ws.getRow(rowcount);
		cell = row.getCell(colcount);
		boolean data;

		try {
			data = cell.getBooleanCellValue();
		} catch (Exception e) {
			data = false;
		}
		wb.close();
		return data;

	}

	public static void setData(String Xlfile, String Xlsheet, int rowcount, int colcount, String data)
			throws IOException {

		fi = new FileInputStream(Xlfile);
		wb = new XSSFWorkbook(fi);
		ws = wb.getSheet(Xlsheet);
		row = ws.getRow(rowcount);
		cell = row.getCell(colcount);
		cell.setCellValue(data);

		fo = new FileOutputStream(Xlfile);
		wb.write(fo);
		wb.close();

	}

	public static void fillGreenColor(String Xlfile, String Xlsheet, int rowcount, int colcount) throws IOException {
		fi = new FileInputStream(Xlfile);
		wb = new XSSFWorkbook(fi);
		ws = wb.getSheet(Xlsheet);
		row = ws.getRow(rowcount);
		cell = row.getCell(colcount);

		style = wb.createCellStyle();
		style.setFillForegroundColor(IndexedColors.BRIGHT_GREEN.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		cell.setCellStyle(style);
		fo = new FileOutputStream(Xlsheet);
		wb.write(fo);
		wb.close();
	}

	public static void fillRedColor(String Xlfile, String Xlsheet, int rowcount, int colcount) throws IOException {
		fi = new FileInputStream(Xlfile);
		wb = new XSSFWorkbook(fi);
		ws = wb.getSheet(Xlsheet);
		row = ws.getRow(rowcount);
		cell = row.getCell(colcount);

		style = wb.createCellStyle();
		style.setFillForegroundColor(IndexedColors.RED.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		cell.setCellStyle(style);

		fo = new FileOutputStream(Xlsheet);
		wb.write(fo);
		wb.close();
	}

}
