package top.onceio.excel.utils;


import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.math.BigDecimal;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelTplWriter {
	private static final Logger LOGGER = LoggerFactory.getLogger(ExcelTplWriter.class);
	public static void  write(String tplPath, int exampleRow, List<Object[]> data, String filepath) {
		FileInputStream fis = null;
		FileOutputStream fos = null;
		try {
			fis = new  FileInputStream(tplPath);
			fos = new  FileOutputStream(filepath);
			write(fis, exampleRow,data, filepath,fos);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
			LOGGER.error(e.getMessage());
		}finally {
			if(fis != null) {
				try {
					fis.close();
				} catch (IOException e) {
					e.printStackTrace();
					LOGGER.error(e.getMessage());
				}
			}
			if(fos != null) {
				try {
					fos.close();
				} catch (IOException e) {
					e.printStackTrace();
					LOGGER.error(e.getMessage());
				}
			}
		}
	}

	public static void  write(InputStream tplis, int exampleRow, List<Object[]> data, String filename, OutputStream os) {
		write(tplis,exampleRow,null,data,filename,os);
	}

	public static void fillCellValue(Cell cell, Object val, String comment) {
		switch(cell.getCellType()) {
			case Cell.CELL_TYPE_NUMERIC:
				if(val instanceof Date) {
					cell.setCellValue((Date)val);
				} else if(val instanceof Long) {
					cell.setCellValue(Double.parseDouble(val.toString()));
				} else {
					cell.setCellValue(Double.parseDouble(val.toString()));
				}
				break;
			case Cell.CELL_TYPE_STRING:
				cell.setCellValue(val.toString());
				break;
			case Cell.CELL_TYPE_BOOLEAN:
				cell.setCellValue(Boolean.parseBoolean(val.toString()));
				break;
			case Cell.CELL_TYPE_FORMULA:
				cell.setCellFormula(val.toString());
				break;
			case Cell.CELL_TYPE_BLANK:
				if(val instanceof String) {
					cell.setCellValue((String)val);
				}else if(val instanceof Boolean) {
					cell.setCellValue((Boolean) val);
				}else if(val instanceof BigDecimal
						|| val instanceof Integer
						|| val instanceof Short
						|| val instanceof Float
						|| val instanceof Double
						|| val instanceof Long) {
					cell.setCellValue(Double.parseDouble(val.toString()));
				}else if(val instanceof Date) {
					cell.setCellValue((Date)val);
				} else {
					cell.setCellValue(val.toString());
				}
				break;
			case Cell.CELL_TYPE_ERROR:
				//cell.setCellErrorValue(val);
				break;
			default:
				LOGGER.info("unresolve : ", cell.getCellType());
		}
	}


	public static void copyCell(Cell cell, int index, Row rowDes) {
		if(cell != null) {
			Cell dest = rowDes.createCell(index);
			dest.setCellType(cell.getCellType());
			dest.setCellStyle(cell.getCellStyle());
			switch (cell.getCellType()) {
				case Cell.CELL_TYPE_BOOLEAN:
					dest.setCellValue(cell.getBooleanCellValue());
					break;
				case Cell.CELL_TYPE_STRING:
					dest.setCellValue(cell.getStringCellValue());
					break;
				case Cell.CELL_TYPE_NUMERIC:
					dest.setCellValue(cell.getNumericCellValue());
					break;
				case Cell.CELL_TYPE_FORMULA:
					dest.setCellValue(cell.getCellFormula());
					break;
				case Cell.CELL_TYPE_BLANK:
					break;
				case Cell.CELL_TYPE_ERROR:
					dest.setCellValue(cell.getErrorCellValue());
					break;
			}
		}
	}

	/**
	 *
	 * @param tplis
	 * @param exampleRow
	 * @param heads 从模板尾部开始替换的头部标题
	 * @param data
	 * @param filename
	 * @param os
	 */
	public static void  write(InputStream tplis, int exampleRow,List<String> heads, List<Object[]> data, String filename, OutputStream os) {
		String ext = filename.substring(filename.lastIndexOf("."));
		Workbook wb = null;
		XSSFWorkbook wbEg = null;
		try {
			if (".xls".equals(ext)) {
				wb = new HSSFWorkbook(tplis);
			} else if (".xlsx".equals(ext)) {
				wbEg = new XSSFWorkbook(tplis);
				wb = new SXSSFWorkbook(wbEg,100);
			} else {
				throw new RuntimeException("上次文件格式不正确（只支持xls和xlsx）");
			}
			Sheet sheetEg = wb.getSheetAt(0);
			int r = exampleRow;
			if(heads != null) {
				Row headRow = sheetEg.getRow(exampleRow-1);
				int offset = headRow.getLastCellNum() - heads.size();
				for(int i=0; i < heads.size(); i++) {
					Cell cell = headRow.getCell(i+offset);
					cell.setCellValue(heads.get(i));
				}
			}
			Row example = sheetEg.getRow(r);
			Map<Integer,Cell> egCell = new HashMap<>();
			for(int c = example.getFirstCellNum(); c <= example.getLastCellNum(); c++) {
				Cell cell = example.getCell(c);
				if(cell != null) {
					egCell.put(c,example.getCell(c));
				}
			}

			Sheet sheet = wb.createSheet();

			for(int i = 0; i < r; i++) {
				Row rowEg = sheetEg.getRow(i);
				Row rowDes = sheet.createRow(i);

				for(int c = rowEg.getFirstCellNum(); c <= rowEg.getLastCellNum(); c++) {
					Cell cell = rowEg.getCell(c);
					copyCell(cell, c, rowDes);
					sheet.setColumnWidth(c, sheetEg.getColumnWidth(c));
				}
			}
			Row row = null;
			for(Object[] objs : data) {
				row = sheet.createRow(r);
				for(int i = 0; i < objs.length; i++) {
					Cell eg = egCell.get(i);
					Cell cell = row.createCell(i);
					Object val = objs[i];
					String comment = null;
					int cellType = Cell.CELL_TYPE_STRING;
					if(eg != null) {
						cell.setCellType(eg.getCellType());
						cell.setCellStyle(eg.getCellStyle());
						Comment cc = eg.getCellComment();
						cellType = cell.getCellType();
						if(cc != null) {
							comment = cc.getString().getString();
						}
					}
					if(eg == null || val == null) continue;
					fillCellValue(cell,val,comment);
				}
				r++;
			}
			wb.removeSheetAt(0);
			wb.write(os);
		} catch (IOException  e) {
			e.printStackTrace();
			LOGGER.error("Exception", e);
		} finally {
			if(wb != null) {
				try {
					wb.close();
				} catch (IOException e) {
					e.printStackTrace();
					LOGGER.error(e.getMessage());
				}
			}
			if(wbEg != null) {
				try {
					wbEg.close();
				} catch (IOException e) {
					e.printStackTrace();
					LOGGER.error(e.getMessage());
				}
			}
		}
	}

	public static void  writeByColumn(InputStream tplis,final int startRow, final int exampleCol, List<Object[]> data, String filename, OutputStream os) {
		String ext = filename.substring(filename.lastIndexOf("."));
		Workbook wb = null;
		Workbook wbEg = null;
		try {
			if (".xls".equals(ext)) {
				wb = new HSSFWorkbook(tplis);
			} else if (".xlsx".equals(ext)) {
				wbEg = new XSSFWorkbook(tplis);
				wb = new SXSSFWorkbook((XSSFWorkbook)wbEg,100);
			} else {
				throw new RuntimeException("上次文件格式不正确（只支持xls和xlsx）");
			}
			Sheet sheetEg = wbEg.getSheetAt(0);
			int r = startRow;
			Map<Integer,Cell> egCell = new HashMap<>();
			for(int c = startRow; c <= sheetEg.getLastRowNum(); c++) {
				Row example = sheetEg.getRow(c);
				Cell cell = example.getCell(exampleCol);
				if(cell != null) {
					egCell.put(c, cell);
				}
			}
			Sheet sheet = wb.createSheet();
			for(int i = 0; i < startRow; i++) {
				Row rowEg = sheetEg.getRow(i);
				if(rowEg != null) {
					Row rowDes = sheet.createRow(i);
					for(int c = rowEg.getFirstCellNum(); c <= rowEg.getLastCellNum(); c++) {
						Cell cell = rowEg.getCell(c);
						copyCell(cell,c,rowDes);
					}
				}
			}
			//复制左侧
			for(int i = startRow; i <= sheetEg.getLastRowNum(); i++) {
				Row example = sheetEg.getRow(i);
				Row rowDes = sheet.createRow(i);
				for(int c = 0; c < exampleCol; c++) {
					Cell cell = example.getCell(c);
					copyCell(cell, c, rowDes);
				}
			}
			Row row = null;
			for(int colNum = 0; colNum < data.size(); colNum++) {
				Object[] objs = data.get(colNum);
				if(objs == null) continue;
				for (int i = 0; i < objs.length; i++) {
					Cell eg = egCell.get(startRow + i);
					row = sheet.getRow(startRow + i);
					Cell cell = row.createCell(colNum + exampleCol);
					cell.setCellStyle(eg.getCellStyle());
					cell.setCellType(eg.getCellType());
					Comment cc = eg.getCellComment();
					String comment = null;
					if (cc != null) {
						comment = cc.getString().getString();
					}
					Object val = objs[i];
					if (val == null) continue;
					fillCellValue(cell,val,comment);
				}
				r++;
			}
			wb.removeSheetAt(0);
			wb.write(os);
		} catch (IOException  e) {
			e.printStackTrace();
			LOGGER.error("Exception", e);
		} finally {
			if(wb != null) {
				try {
					wb.close();
				} catch (IOException e) {
					LOGGER.error(e.getMessage());
				}
			}
			if(wbEg != null) {
				try {
					wbEg.close();
				} catch (IOException e) {
					LOGGER.error(e.getMessage());
				}
			}
		}
	}

	private static Object parseNumber(double val, Class<?> type) {
		if (type.equals(long.class) || type.equals(Long.class)) {
			return (long) val;
		} else if (type.equals(int.class) || type.equals(Integer.class)) {
			return (int) val;
		} else if (type.equals(short.class) || type.equals(Short.class)) {
			return (long) val;
		} else if (type.equals(double.class) || type.equals(Double.class)) {
			return val;
		} else if (type.equals(float.class) || type.equals(Float.class)) {
			return (float) val;
		} else if (type.equals(byte.class) || type.equals(Byte.class)) {
			return (byte) val;
		} else if (type.equals(BigDecimal.class)) {
			return new BigDecimal(val);
		}

		return null;
	}

	@SuppressWarnings("unchecked")
	private static <T> T strToBaseType(Class<T> type, String val) {
		if (val == null) {
			return null;
		} else if (type.equals(String.class)) {
			return (T) val;
		} else if (type.equals(int.class) || type.equals(Integer.class)) {
			return (T) Integer.valueOf(val);
		} else if (type.equals(long.class) || type.equals(Long.class)) {
			return (T) Long.valueOf(val);
		} else if (type.equals(boolean.class) || type.equals(Boolean.class)) {
			return (T) Boolean.valueOf(val);
		} else if (type.equals(byte.class) || type.equals(Byte.class)) {
			return (T) Byte.valueOf(val);
		} else if (type.equals(short.class) || type.equals(Short.class)) {
			return (T) Short.valueOf(val);
		} else if (type.equals(double.class) || type.equals(Double.class)) {
			return (T) Double.valueOf(val);
		} else if (type.equals(float.class) || type.equals(Float.class)) {
			return (T) Float.valueOf(val);
		} else if (type.equals(BigDecimal.class)) {
			return (T) BigDecimal.valueOf(Double.valueOf(val));
		}
		return null;
	}

}
