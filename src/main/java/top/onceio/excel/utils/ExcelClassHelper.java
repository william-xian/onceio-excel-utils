package top.onceio.excel.utils;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.util.*;

public class ExcelClassHelper {
	private static final Logger LOGGER = LoggerFactory.getLogger(ExcelClassHelper.class);

	private static Map<String, Field> getNameToField(Class<?> clazz, Map<String, String> aliasToField) {
		Set<String> fieldNames = null;
		if (aliasToField != null) {
			fieldNames = new HashSet<>(aliasToField.values());
		}
		Map<String, Field> nameToField = new HashMap<>();
		for (Class<?> sc = clazz; sc != null && !sc.equals(Object.class); sc = sc.getSuperclass()) {
			for (Field f : sc.getDeclaredFields()) {
				if (fieldNames == null) {
					nameToField.put(f.getName(), f);
				} else if (fieldNames.contains(f.getName())) {
					nameToField.put(f.getName(), f);
				}
			}
		}
		return nameToField;
	}

	private static Map<Field, Integer> build(Class<?> clazz, Map<String, String> aliasToField, Map<String, Integer> nameToIndex) {
		Map<String, Field> nameToField = getNameToField(clazz, aliasToField);
		Map<Field, Integer> fieldToIndex = new HashMap<>();
		for (Map.Entry<String, Integer> entry : nameToIndex.entrySet()) {
			String fn = entry.getKey();
			if (aliasToField != null) {
				fn = aliasToField.get(entry.getKey());
				if (fn == null) {
					fn = entry.getKey();
				}
			}
			if (fn != null) {
				Field f = nameToField.get(fn);
				if (f != null) {
					f.setAccessible(true);
					fieldToIndex.put(f, entry.getValue());
				}
			}
		}
		return fieldToIndex;
	}

	public static <T> List<T> read(Class<T> clazz, Map<String, String> alias, String filepath) {
		List<T> result = null;
		FileInputStream fis = null;
		try {
			fis = new FileInputStream(filepath);
			result = read(clazz, alias, filepath, fis);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
			LOGGER.error(e.getMessage());
		} finally {
			if (fis != null) {
				try {
					fis.close();
				} catch (IOException e) {
					e.printStackTrace();
					LOGGER.error(e.getMessage());
				}
			}
		}
		return result;
	}

	public static <T> List<T> read(Class<T> clazz, Map<String, String> alias, String filename, InputStream is) {
		return read(clazz, alias, filename, is, true);
	}

	private static void fillObjectViaExcel(int valType, Cell cell, Field field, Object obj, boolean useTrim) throws IllegalAccessException {
		switch (valType) {
			case Cell.CELL_TYPE_NUMERIC:
				if (field.getType().equals(Date.class)) {
					field.set(obj, cell.getDateCellValue());
				} else if (field.getType().equals(Long.class)
						&& (cell.getNumericCellValue() - (long) cell.getNumericCellValue()) != 0.0) {
					field.set(obj, cell.getDateCellValue().getTime());
				} else {
					field.set(obj, parseNumber(cell.getNumericCellValue(), field.getType()));
				}
				break;
			case Cell.CELL_TYPE_STRING:
				if (cell.getStringCellValue() != null) {
					if (useTrim) {
						field.set(obj, strToBaseType(field.getType(), cell.getStringCellValue().trim()));
					} else {
						field.set(obj, strToBaseType(field.getType(), cell.getStringCellValue()));
					}
				}
				break;
			case Cell.CELL_TYPE_BOOLEAN:
				field.set(obj, cell.getBooleanCellValue());
				break;
			default:
		}
	}

	public static <T> List<T> read(Class<T> clazz, Map<String, String> alias, String filename, InputStream is, boolean useTrim) {
		List<T> result = new ArrayList<>();
		String ext = filename.substring(filename.lastIndexOf("."));
		Workbook wb = null;
		int rowNum = 0;
		int colNum = 0;
		Cell cell = null;
		try {
			if (".xls".equals(ext)) {
				wb = new HSSFWorkbook(is);
			} else if (".xlsx".equals(ext)) {
				wb = new XSSFWorkbook(is);
			} else {
				throw new RuntimeException("文件格式只支持xls和xlsx");
			}

			Sheet sheet = wb.getSheetAt(0);
			Map<String, Integer> nameToC = new HashMap<>();
			Row row = sheet.getRow(sheet.getFirstRowNum());
			for (int c = row.getFirstCellNum(); c <= row.getLastCellNum(); c++) {
				cell = row.getCell(c);
				if (cell != null) {
					if (useTrim && cell.getStringCellValue() != null) {
						nameToC.put(cell.getStringCellValue().trim(), c);
					} else {
						nameToC.put(cell.getStringCellValue(), c);
					}
				}
			}
			Map<Field, Integer> fieldToIndex = build(clazz, alias, nameToC);
			for (rowNum = sheet.getFirstRowNum() + 1; rowNum <= sheet.getLastRowNum(); rowNum++) {
				row = sheet.getRow(rowNum);
				if (row == null) continue;
				T obj = clazz.newInstance();
				for (Map.Entry<Field, Integer> entry : fieldToIndex.entrySet()) {
					colNum = entry.getValue();
					cell = row.getCell(entry.getValue());
					Field field = entry.getKey();
					if (cell == null) continue;
					switch (cell.getCellType()) {
						case Cell.CELL_TYPE_NUMERIC:
							fillObjectViaExcel(Cell.CELL_TYPE_NUMERIC, cell, field, obj, useTrim);
							break;
						case Cell.CELL_TYPE_STRING:
							fillObjectViaExcel(Cell.CELL_TYPE_STRING, cell, field, obj, useTrim);
							break;
						case Cell.CELL_TYPE_BOOLEAN:
							fillObjectViaExcel(Cell.CELL_TYPE_BOOLEAN, cell, field, obj, useTrim);
							break;
						case Cell.CELL_TYPE_FORMULA:
							do {
								try {
									if (field.getType().equals(Date.class)) {
										field.set(obj, cell.getDateCellValue());
										break;
									} else if (field.getType().equals(Long.class)
											&& (cell.getNumericCellValue() - (long) cell.getNumericCellValue()) != 0.0) {
										field.set(obj, cell.getDateCellValue().getTime());
										break;
									} else if (field.getType().equals(String.class)) {
										field.set(obj, cell.getStringCellValue());
										break;
									} else if (field.getType().equals(Boolean.class)) {
										field.set(obj, cell.getBooleanCellValue());
										break;
									} else if (field.getType().equals(Integer.class)
											|| field.getType().equals(Short.class)
											|| field.getType().equals(Byte.class)
											|| field.getType().equals(Float.class)
											|| field.getType().equals(BigDecimal.class)) {
										field.set(obj, parseNumber(cell.getNumericCellValue(), field.getType()));
										break;
									} else {
										field.set(obj, strToBaseType(field.getType(), cell.getStringCellValue()));
										break;
									}
								} catch (Exception formulaE) {
								}
								//TODO 容错处理
								try {
									Double fdVal = cell.getNumericCellValue();
									if (fdVal != null) {
										fillObjectViaExcel(Cell.CELL_TYPE_NUMERIC, cell, field, obj, useTrim);
										break;
									}
								} catch (Exception formulaE) {
								}
								try {
									String fsVal = cell.getStringCellValue();
									if (fsVal != null) {
										fillObjectViaExcel(Cell.CELL_TYPE_STRING, cell, field, obj, useTrim);
										break;
									}
								} catch (Exception formulaE) {
								}
								try {
									Boolean fbVal = cell.getBooleanCellValue();
									if (fbVal != null) {
										fillObjectViaExcel(Cell.CELL_TYPE_BOOLEAN, cell, field, obj, useTrim);
										break;
									}
								} catch (Exception formulaE) {
								}
							} while (false);
							break;
						case Cell.CELL_TYPE_BLANK:
							break;
						case Cell.CELL_TYPE_ERROR:
							break;
						default:
							LOGGER.info("未知类型 : ", cell.getCellType());
					}
				}
				result.add(obj);
			}

		} catch (InstantiationException | IllegalAccessException | IOException e) {
			e.printStackTrace();
			LOGGER.error("Exception", e);
		} catch (IllegalStateException ex) {
			ex.printStackTrace();
			LOGGER.error("数据错误:" + String.format("%s,%s", rowNum, colNum));
		} finally {
			if (wb != null) {
				try {
					wb.close();
				} catch (IOException e) {
					e.printStackTrace();
					LOGGER.error(e.getMessage());
				}
			}
		}
		return result;
	}


	public static <T> void write(Class<T> clazz, List<T> data, Map<String, String> alias, String tplPath, String filepath) {
		FileInputStream fis = null;
		FileOutputStream fos = null;

		try {
			fis = new FileInputStream(tplPath);
			fos = new FileOutputStream(filepath);
			write(clazz, data, alias, filepath, fis, fos);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
			LOGGER.error(e.getMessage());
		} finally {
			if (fis != null) {
				try {
					fis.close();
				} catch (IOException e) {
					e.printStackTrace();
					LOGGER.error(e.getMessage());
				}
			}
			if (fos != null) {
				try {
					fos.close();
				} catch (IOException e) {
					e.printStackTrace();
					LOGGER.error(e.getMessage());
				}
			}
		}
	}


	public static <T> void write(Class<T> clazz, List<T> data, Map<String, String> alias, String filename, InputStream tplis, OutputStream os) {
		String ext = filename.substring(filename.lastIndexOf("."));
		Workbook wb = null;

		if (tplis == null) {
			throw new RuntimeException("模板不可为null");
		}
		if (os == null) {
			throw new RuntimeException("输出流不可为空");
		}
		try {
			if (".xls".equals(ext)) {
				wb = new HSSFWorkbook(tplis);
			} else if (".xlsx".equals(ext)) {
				wb = new XSSFWorkbook(tplis);
			} else {
				throw new RuntimeException("上次文件格式不正确（只支持xls和xlsx）");
			}

			Sheet sheet = wb.getSheetAt(0);
			Map<String, Integer> nameToC = new HashMap<>();
			Row row = sheet.getRow(sheet.getFirstRowNum());
			int r = sheet.getFirstRowNum() + 1;
			Row example = sheet.getRow(r++);
			Map<Integer, Cell> egCell = new HashMap<>();
			for (int c = row.getFirstCellNum(); c <= row.getLastCellNum(); c++) {
				Cell cell = row.getCell(c);
				if (cell != null) {
					nameToC.put(cell.getStringCellValue(), c);
					egCell.put(c, example.getCell(c));
				}

			}
			Map<Field, Integer> fieldToIndex = build(clazz, alias, nameToC);
			boolean isEg = true;
			for (T obj : data) {
				if (isEg) {
					row = example;
				} else {
					row = sheet.createRow(r++);
				}
				for (Map.Entry<Field, Integer> entry : fieldToIndex.entrySet()) {
					Cell eg = egCell.get(entry.getValue());
					Field field = entry.getKey();
					if (eg == null || field == null) continue;
					Cell cell = null;
					if(!isEg) {
						cell = row.createCell(entry.getValue());
						cell.setCellStyle(eg.getCellStyle());
						cell.setCellType(eg.getCellType());
						Comment cc = eg.getCellComment();
						if (cc != null) {
							cell.setCellComment(cc);
						}
					} else {
						cell = row.getCell(entry.getValue());
					}
					Object val = field.get(obj);
					if (val == null) continue;
					fillCellValue(cell, val);
				}

				isEg = false;
			}
			wb.write(os);
		} catch (IllegalAccessException | IOException e) {
			e.printStackTrace();
			LOGGER.error("Exception", e);
		} catch (IllegalStateException e) {
			e.printStackTrace();
			LOGGER.error("Exception", e);
		} finally {
			if (wb != null) {
				try {
					wb.close();
				} catch (IOException e) {
					e.printStackTrace();
					LOGGER.error(e.getMessage());
				}
			}
		}
	}

	public static void fillCellValue(Cell cell, Object val) {
		switch (cell.getCellType()) {
			case Cell.CELL_TYPE_NUMERIC:
				if (val instanceof Date) {
					cell.setCellValue((Date) val);
				} else if (val instanceof Long) {
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
				if (val instanceof String) {
					cell.setCellValue((String) val);
				} else if (val instanceof Boolean) {
					cell.setCellValue((Boolean) val);
				} else if (val instanceof BigDecimal
						|| val instanceof Integer
						|| val instanceof Short
						|| val instanceof Float
						|| val instanceof Double
						|| val instanceof Long) {
					cell.setCellValue(Double.parseDouble(val.toString()));
				} else if (val instanceof Date) {
					cell.setCellValue((Date) val);
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
			return new BigDecimal(val + "");
		} else if (type.equals(String.class)) {
			if (val - (long) val == 0.0) {
				return (long) val + "";
			} else {
				return val + "";
			}
		}

		return null;
	}

	@SuppressWarnings("unchecked")
	private static <T> T strToBaseType(Class<T> type, String val) {
		if (val != null) {
			if (type.equals(String.class)) {
				return (T) val;
			} else if (val.trim().equals("")) {
				return null;
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
		}
		return null;
	}

}
