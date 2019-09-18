package com.test.excel.poi;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * Excel导入导出
 */
public class ExcelUtil {

	public static final String TEMP_FILE_NAME = "\\temp_files\\";
	public static final String LOG_FILE_NAME = "\\log_files\\";

	// begin 导入导出后缀.xls

	/**
	 * 导出Excel【.xls】
	 * 
	 * @param title
	 *            表格的名称
	 * 
	 * @param headersName
	 *            表格头【中文】
	 * 
	 * @param headersId
	 *            表格头【英文】
	 * 
	 * @param listB
	 *            表格头
	 * 
	 * @param filePath
	 *            文件保存路径【例如："d://ceshi.xls",注意：后缀必须是.xls】
	 */
	private static int exportExcel03(String title, List headersName, List headersId, List<Map> listB, String filePath) {
		/* （一）表头--标题栏 */
		Map<Integer, String> headersNameMap = new HashMap<>();
		int key = 0;
		for (int i = 0; i < headersName.size(); i++) {
			if (!headersName.get(i).equals(null)) {
				headersNameMap.put(key, headersName.get(i).toString());
				key++;
			}
		}
		/* （二）字段 */
		Map<Integer, String> titleFieldMap = new HashMap<>();
		int value = 0;
		for (int i = 0; i < headersId.size(); i++) {
			if (!headersId.get(i).equals(null)) {
				titleFieldMap.put(value, headersId.get(i).toString());
				value++;
			}
		}
		/* （三）声明一个工作薄：包括构建工作簿、表格、样式 */
		HSSFWorkbook wb = new HSSFWorkbook();
		HSSFSheet sheet = wb.createSheet(title);
		sheet.setDefaultColumnWidth((short) 15);
		// 生成一个样式
		HSSFCellStyle style = wb.createCellStyle();
		HSSFRow row = sheet.createRow(0);
		style.setAlignment(HorizontalAlignment.CENTER);
		HSSFCell cell;
		Collection<String> c = headersNameMap.values();// 拿到表格所有标题的value的集合
		Iterator<String> it = c.iterator();// 表格标题的迭代器
		/* （四）导出数据：包括导出标题栏以及内容栏 */
		// 根据选择的字段生成表头
		short size = 0;
		while (it.hasNext()) {
			cell = row.createCell(size);
			cell.setCellValue(it.next().toString());
			cell.setCellStyle(style);
			size++;
		}

		int zdRow = 1;// 真正的数据记录的列序号
		for (int i = 0; i < listB.size(); i++) {
			Map<String, Object> mapTemp = listB.get(i);
			row = sheet.createRow(zdRow);
			zdRow++;
			// Set<String> keys = mapTemp.keySet();
			// Iterator<String> iters = keys.iterator();
			for (int j = 0; j < headersName.size(); j++) {
				String tempField = headersId.get(j).toString();
				if(mapTemp.get(tempField) == null) {
					row.createCell(j).setCellValue("");
				} else {
					row.createCell(j).setCellValue(mapTemp.get(tempField).toString());
				}
			}
		}
		try {
			File parent = new File(filePath).getParentFile();
			if (!parent.exists()) {
				parent.mkdir();
			}
			FileOutputStream exportXls = new FileOutputStream(filePath);
			wb.write(exportXls);
			exportXls.close();
			wb.close();
			return 0;
		} catch (FileNotFoundException e) {
			e.printStackTrace();
			return 10003;
		} catch (IOException e) {
			e.printStackTrace();
			return 20004;
		}
	}

	/**
	 * 导入Excel
	 * 
	 * @param filePath
	 *            文件保存路径【例如："d://ceshi.xls",注意：后缀必须是.xls】
	 */
	private static List<HashMap<String, Object>> readExcel03(String filePath) throws IOException {
		List<HashMap<String, Object>> list = new ArrayList<>();
		FileInputStream inputStream = new FileInputStream(new File(filePath));
		// 读取工作簿
		HSSFWorkbook workBook = new HSSFWorkbook(inputStream);
		// 读取工作表
		HSSFSheet sheet = workBook.getSheetAt(0);
		if (sheet.getLastRowNum() > 0) {
			HSSFRow title_row = sheet.getRow(0);
			for (int t = 1; t <= sheet.getLastRowNum() && sheet.getRow(t) != null; t++) {
				int cellNums = sheet.getRow(t).getLastCellNum();
				HSSFRow row = sheet.getRow(t);
				LinkedHashMap<String, Object> map = new LinkedHashMap<>();
				for (int i = 0; i < cellNums; i++) {
					if (null == row.getCell(i)) {
						map.put(title_row.getCell(i).getStringCellValue(), "");
						continue;
					}
					CellType cellType = row.getCell(i).getCellTypeEnum();
					// NUMERIC格式的需要判断一下
					if (cellType == CellType.NUMERIC) {
						if(HSSFDateUtil.isCellDateFormatted(row.getCell(i))){
							SimpleDateFormat df = new SimpleDateFormat("YYYY-MM-dd HH:mm:ss");
							map.put(title_row.getCell(i).getStringCellValue(), df.format(row.getCell(i).getDateCellValue()));
						} else {
							String value = Helper.doubleToString(row.getCell(i).getNumericCellValue());
							map.put(title_row.getCell(i).getStringCellValue(), value);
						}
					} else if (cellType == CellType.STRING)  {
						map.put(title_row.getCell(i).getStringCellValue(), row.getCell(i).getStringCellValue());
					} else if (cellType == CellType.FORMULA) {
					    String value = Helper.doubleToString(row.getCell(i).getNumericCellValue());
                        map.put(title_row.getCell(i).getStringCellValue(), value);
					}
				}
				list.add(map);
			}
			inputStream.close();
			workBook.close();
		}
		return list;
	}
	// end

	// begin 导入导出后缀.xlsx

	/**
	 * 导出Excel【.xlsx】
	 * 
	 * @param title
	 *            表格的名称
	 * 
	 * @param headersName
	 *            表格头【中文】
	 * 
	 * @param headersId
	 *            表格头【英文】
	 * 
	 * @param listB
	 *            表格头
	 * 
	 * @param filePath
	 *            文件保存路径【例如："d://ceshi.xlsx",注意：后缀必须是.xlsx】
	 */
	private static int exportExcel07(String title, List headersName, List headersId, List<Map> listB, String filePath) {

		/* （一）表头--标题栏 */
		Map<Integer, String> headersNameMap = new HashMap<>();
		int key = 0;
		for (int i = 0; i < headersName.size(); i++) {
			if (!headersName.get(i).equals(null)) {
				headersNameMap.put(key, headersName.get(i).toString());
				key++;
			}
		}
		/* （二）字段 */
		Map<Integer, String> titleFieldMap = new HashMap<>();
		int value = 0;
		for (int i = 0; i < headersId.size(); i++) {
			if (!headersId.get(i).equals(null)) {
				titleFieldMap.put(value, headersId.get(i).toString());
				value++;
			}
		}

		/*
		 * （三）声明一个工作薄：包括构建工作簿、表格、样式
		 */
		// 创建工作簿
		XSSFWorkbook wb = new XSSFWorkbook();
		// 创建工作表 工作表的名字叫 title
		XSSFSheet sheet = wb.createSheet(title);
		sheet.setDefaultColumnWidth((short) 15);
		// 表头样式
		XSSFCellStyle title_style = wb.createCellStyle();
		title_style.setAlignment(HorizontalAlignment.CENTER);
		XSSFRow row = sheet.createRow(0);
		XSSFCell cell;
		/*
		 * （四）导出数据：包括导出标题栏以及内容栏
		 */
		// 根据选择的字段生成表头--标题
		Collection<String> c = headersNameMap.values();
		Iterator<String> it = c.iterator();
		short size = 0;
		while (it.hasNext()) {
			cell = row.createCell(size);
			cell.setCellValue(it.next());
			cell.setCellStyle(title_style);
			size++;
		}

		int zdRow = 1;// 真正的数据记录的列序号
		for (int i = 0; i < listB.size(); i++) {
			Map<String, Object> mapTemp = listB.get(i);
			row = sheet.createRow(zdRow);
			zdRow++;
			for (int j = 0; j < headersId.size(); j++) {
				String tempField = headersId.get(j).toString();
				if(mapTemp.get(tempField) == null) {
					row.createCell(j).setCellValue("");
				} else {
					row.createCell(j).setCellValue(mapTemp.get(tempField).toString());
				}
			}
		}
		try {
			File parent = new File(filePath).getParentFile();
			if (!parent.exists()) {
				parent.mkdir();
			}
			FileOutputStream exportXls = new FileOutputStream(filePath);
			wb.write(exportXls);
			exportXls.close();
			wb.close();
			return 0;
		} catch (FileNotFoundException e) {
			e.printStackTrace();
			return 10003;
		} catch (IOException e) {
			e.printStackTrace();
			return 20004;
		}
	}

	/**
	 * 导入Excel
	 * 
	 * @param filePath
	 *            文件保存路径【例如："d://ceshi.xlsx",注意：后缀必须是.xlsx】
	 */
	private static List<HashMap<String, Object>> readExcel07(String filePath) throws IOException {
		List<HashMap<String, Object>> list = new ArrayList<>();
		FileInputStream inputStream = new FileInputStream(new File(filePath));
		// 读取工作簿
		XSSFWorkbook workBook = new XSSFWorkbook(inputStream);
		// 读取工作表
		XSSFSheet sheet = workBook.getSheetAt(0);
		if (sheet.getLastRowNum() > 0) {
			XSSFRow title_row = sheet.getRow(0);
			for (int t = 1; t <= sheet.getLastRowNum() && sheet.getRow(t) != null; t++) {
				int cellNums = sheet.getRow(t).getLastCellNum();
				XSSFRow row = sheet.getRow(t);
				LinkedHashMap<String, Object> map = new LinkedHashMap<>();
				for (int i = 0; i < cellNums; i++) {
					if (null == row.getCell(i)) {
						map.put(title_row.getCell(i).getStringCellValue(), "");
						continue;
					}
					CellType cellType = row.getCell(i).getCellTypeEnum();
//					 NUMERIC格式的需要判断一下
					if (cellType == CellType.NUMERIC) {
						if(HSSFDateUtil.isCellDateFormatted(row.getCell(i))){
							SimpleDateFormat df = new SimpleDateFormat("YYYY-MM-dd HH:mm:ss");
							map.put(title_row.getCell(i).getStringCellValue(), df.format(row.getCell(i).getDateCellValue()));
						} else {
							String value = Helper.doubleToString(row.getCell(i).getNumericCellValue());
							map.put(title_row.getCell(i).getStringCellValue(), value);
						}
						
					} else if (cellType == CellType.STRING)  {
						map.put(title_row.getCell(i).getStringCellValue(), row.getCell(i).getStringCellValue());
					} else if (cellType == CellType.FORMULA) {
					    String value = Helper.doubleToString(row.getCell(i).getNumericCellValue());
                        map.put(title_row.getCell(i).getStringCellValue(), value);
					}
				}
				list.add(map);
			}
			inputStream.close();
			workBook.close();
		}
		return list;
	}
	// end

	// begin Excel导入导出
	/**
	 * 导出Excel
	 * 
	 * @param title
	 *            表格的名称
	 * 
	 * @param headersName
	 *            表格头【中文】
	 * 
	 * @param headersId
	 *            表格头【英文】
	 * 
	 * @param listB
	 *            表格头
	 * 
	 * @param filePath
	 *            文件保存路径
	 */
	public static int exportExcel(String title, List headersName, List headersId, List<Map> listB, String filePath) {
		int resule = -1;
		String ext = getExtName(filePath);
		if (ext.equals(".xls")) {
			resule = exportExcel03(title, headersName, headersId, listB, filePath);
		} else if (ext.equals(".xlsx")) {
			resule = exportExcel07(title, headersName, headersId, listB, filePath);
		}
		return resule;
	}

	/**
	 * 导出Excel
	 * 
	 * @param title
	 *            表格的名称
	 * 
	 * @param headersName
	 *            表格头【中文】
	 * 
	 * @param headersId
	 *            表格头【英文】
	 * 
	 * @param listB
	 *            表格头
	 * 
	 * @param filePath
	 *            文件保存路径
	 */
	public static int exportExcel(String title, List headersName, List headersId, List<Map<String, Object>> listB,
			String filePath, int not) {
		List<Map> list = new ArrayList();
		for (int i = 0; i < listB.size(); i++) {
			list.add(listB.get(i));
		}
		int resule = -1;
		String ext = getExtName(filePath);
		if (ext.equals(".xls")) {
			resule = exportExcel03(title, headersName, headersId, list, filePath);
		} else if (ext.equals(".xlsx")) {
			resule = exportExcel07(title, headersName, headersId, list, filePath);
		}
		return resule;
	}

	/**
	 * 导入Excel
	 * 
	 * @param filePath
	 *            文件保存路径
	 */
	public static List<HashMap<String, Object>> readExcel(String filePath) throws IOException {
		List<HashMap<String, Object>> resule = new ArrayList<>();
		String ext = getExtName(filePath);
		if (ext.equals(".xls")) {
			resule = readExcel03(filePath);
		} else if (ext.equals(".xlsx")) {
			resule = readExcel07(filePath);
		}
		return resule;
	}

	private static String getExtName(String path) {
		int index = path.lastIndexOf(".");
		int len = path.length();
		String res = (index > 0 ? (index + 1) == len ? " " : path.substring(index, len) : " ");
		return res;
	}

	// end

	public static void main(String[] args) throws IOException {
//		 daochu();
		try {
			Thread.sleep(10000);
		} catch (InterruptedException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		daoru();
	}

	public static void daochu() {
		List<String> listName = new ArrayList<>();
		listName.add("名字");
		listName.add("编号");
		listName.add("性别");
		List<String> listId = new ArrayList<>();
		listId.add("name");
		listId.add("id");
		listId.add("sex");
		List<Map> listB = new ArrayList<>();
		for (int t = 0; t < 3; t++) {
			Map<String, Object> map = new HashMap<String, Object>();
			map.put("id", "编号s" + t);
			map.put("name", "姓名" + t);
			map.put("sex", "男" + t);
			listB.add(map);
		}
		System.out.println("listB  : " + listB.toString());
		System.out.println(ExcelUtil.exportExcel("测试", listName, listId, listB, "c://test/abcdef.xlsx"));
	}

	public static void daoru() throws IOException {
		// ExcelUtil exportExcelUtil = new ExcelUtil();
		List<HashMap<String, Object>> list = ExcelUtil.readExcel("c://test/abcdef.xlsx");
		System.out.println(list);
	}
	
	/**
	 * 这个方法是生成有两行表头的excel的方法，包含了表头的合并行 	导出格式为 .xlsx
	 * @param List<String>headerName1 第一行表头，需要合并的行要重复列举出来
	 * @param List<String>headnum1 
	 *  每项的形式 "x,x,x,x" 从左到右依次是 起始行，终止行，起始列，终止列，
	 *  总长度 headerName1去重后对应的长度
	 * @param List<String>headerName2 第二行标题的headerName
	 * @param List<String>headerNames 数据对应字段的表头
	 * @param List<String> headerIds 数据对应的字段,
	 * @param List<Map> data 需要写入的数据
	 * @param String filePath 文件地址
	 * @param int title2_start 第二行表头开始的位置
	 * @param fileName 文件名称
	 * */
	public static int exportExcelWithDoubleTitle07(List<String>headerName1, List<String>headnum1,
	        List<String>headerName2,List<String>headerNames,List<String> headerIds,List<Map> data,
	        String filePath,int title2_start,String sheetName) {
		XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet(sheetName);
        XSSFCellStyle style = wb.createCellStyle(); 
        style.setFillForegroundColor(HSSFColor.SKY_BLUE.index);
        style.setAlignment(HorizontalAlignment.CENTER);
        XSSFFont font = wb.createFont();
        
        // 设置字体
        font.setFontName("微软雅黑");
        // 设置字体大小
        font.setFontHeightInPoints((short) 16);
        font.setColor(HSSFColor.ROSE.index);
        font.setBold(true);
        style.setFont(font);
        XSSFCellStyle style2 = wb.createCellStyle();
        style2.setFillForegroundColor(HSSFColor.LIGHT_YELLOW.index);
        XSSFFont font2 = wb.createFont();
        // 设置字体
        font2.setFontName("微软雅黑");
        // 设置字体大小
        font2.setFontHeightInPoints((short) 10);
        style2.setFont(font2);
        XSSFRow row = sheet.createRow(0);
        for (int i = 0 ; i < headerName1.size(); i++) {
//            sheet.autoSizeColumn(i,true);
            sheet.setDefaultColumnWidth((short) 15);
            XSSFCell cell = row.createCell(i);
            cell.setCellValue(headerName1.get(i));
//            if (i > 1) {
                cell.setCellStyle(style);    
//            }
            
        }
	   	//合并单元格
        for (int i = 0 ; i < headnum1.size(); i++) {
//            sheet.autoSizeColumn(i, true);
            String[] temp = headnum1.get(i).split(",");
            Integer startrow = Integer.parseInt(temp[0]);
            Integer overrow = Integer.parseInt(temp[1]);
            Integer startcol = Integer.parseInt(temp[2]);
            Integer overcol = Integer.parseInt(temp[3]);
            sheet.addMergedRegion(new CellRangeAddress(startrow,overrow,startcol,overcol));
        }
        row = row = sheet.createRow(1);
        for (int i = 0; i < headerName2.size();i++) {
//            sheet.autoSizeColumn(i,true);
            sheet.setDefaultColumnWidth((short) 15);
            XSSFCell cell = row.createCell(i);
            cell.setCellValue(headerName2.get(i));
            cell.setCellStyle(style2);
        }
        int tureRow = 2;
        for (int i = 0 ; i < data.size(); i++) {
            Map<String, String> mapTemp = data.get(i);
            row = sheet.createRow(tureRow);
            tureRow ++;
            for (int j = 0 ; j < headerNames.size(); j++) {
                String tempField = headerIds.get(j).toString();
                if (mapTemp.get(tempField) != null) {
//                    sheet.autoSizeColumn(i,true);
                    sheet.setDefaultColumnWidth((short) 15);
                    row.createCell(j).setCellValue(String.valueOf(mapTemp.get(tempField)));// 写进excel对象
                }
            }
            
        }
        try {
            File parent = new File(filePath).getParentFile();
            if (!parent.exists()) {
                parent.mkdir();
            }
            FileOutputStream exportXls = new FileOutputStream(filePath);
            wb.write(exportXls);
            exportXls.close();
            wb.close();
            return 0;
        } catch (FileNotFoundException e) {
            e.printStackTrace();
            return 10003;
        } catch (IOException e) {
            e.printStackTrace();
            return 20004;
        }
	}
	
	/**
	 * 这个方法是生成有两行表头的excel的方法，包含了表头的合并行 	导出格式为 .xls
	 * @param List<String>headerName1 第一行表头，需要合并的行要重复列举出来
	 * @param List<String>headnum1 
	 *  每项的形式 "x,x,x,x" 从左到右依次是 起始行，终止行，起始列，终止列，
	 *  总长度 headerName1去重后对应的长度
	 * @param List<String>headerName2 第二行标题的headerName
	 * @param List<String>headerNames 数据对应字段的表头
	 * @param List<String> headerIds 数据对应的字段,
	 * @param List<Map> data 需要写入的数据
	 * @param String filePath 文件地址
	 * @param int title2_start 第二行表头开始的位置
	 * @param fileName 文件名称
	 * */
	public static int exportExcelWithDoubleTitle03(List<String>headerName1, List<String>headnum1,
	        List<String>headerName2,List<String>headerNames,List<String> headerIds,List<Map> data,
	        String filePath,int title2_start,String sheetName) {
	    HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.createSheet(sheetName);
        HSSFCellStyle style = wb.createCellStyle(); 
        style.setFillForegroundColor(HSSFColor.SKY_BLUE.index);
        style.setAlignment(HorizontalAlignment.CENTER);
        HSSFFont font = wb.createFont();
        
        // 设置字体
        font.setFontName("微软雅黑");
        // 设置字体大小
        font.setFontHeightInPoints((short) 16);
        font.setColor(HSSFColor.ROSE.index);
        font.setBold(true);
        style.setFont(font);
        HSSFCellStyle style2 = wb.createCellStyle();
        style2.setFillForegroundColor(HSSFColor.LIGHT_YELLOW.index);
        HSSFFont font2 = wb.createFont();
        // 设置字体
        font2.setFontName("微软雅黑");
        // 设置字体大小
        font2.setFontHeightInPoints((short) 10);
        style2.setFont(font2);
        HSSFRow row = sheet.createRow(0);
        for (int i = 0 ; i < headerName1.size(); i++) {
//            sheet.autoSizeColumn(i,true);
            sheet.setDefaultColumnWidth((short) 15);
            HSSFCell cell = row.createCell(i);
            cell.setCellValue(headerName1.get(i));
//            if (i > 1) {
                cell.setCellStyle(style);    
//            }
            
        }
	   	//合并单元格
        for (int i = 0 ; i < headnum1.size(); i++) {
//            sheet.autoSizeColumn(i, true);
            String[] temp = headnum1.get(i).split(",");
            Integer startrow = Integer.parseInt(temp[0]);
            Integer overrow = Integer.parseInt(temp[1]);
            Integer startcol = Integer.parseInt(temp[2]);
            Integer overcol = Integer.parseInt(temp[3]);
            sheet.addMergedRegion(new CellRangeAddress(startrow,overrow,startcol,overcol));
        }
        row = row = sheet.createRow(1);
        for (int i = 0; i < headerName2.size();i++) {
//            sheet.autoSizeColumn(i,true);
            sheet.setDefaultColumnWidth((short) 15);
            HSSFCell cell = row.createCell(i);
            cell.setCellValue(headerName2.get(i));
            cell.setCellStyle(style2);
        }
        int tureRow = 2;
        for (int i = 0 ; i < data.size(); i++) {
            Map<String, String> mapTemp = data.get(i);
            row = sheet.createRow(tureRow);
            tureRow ++;
            for (int j = 0 ; j < headerNames.size(); j++) {
                String tempField = headerIds.get(j).toString();
                if (mapTemp.get(tempField) != null) {
//                    sheet.autoSizeColumn(i,true);
                    sheet.setDefaultColumnWidth((short) 15);
                    row.createCell(j).setCellValue(String.valueOf(mapTemp.get(tempField)));// 写进excel对象
                }
            }
            
        }
        try {
            File parent = new File(filePath).getParentFile();
            if (!parent.exists()) {
                parent.mkdir();
            }
            FileOutputStream exportXls = new FileOutputStream(filePath);
            wb.write(exportXls);
            exportXls.close();
            wb.close();
            return 0;
        } catch (FileNotFoundException e) {
            e.printStackTrace();
            return 10003;
        } catch (IOException e) {
            e.printStackTrace();
            return 20004;
        }
	}
	
	
	
	
	
	/**
	 * 这个方法是读取excel（双表头）的数据的方法
	 * @param String filePath 文件路径
	 * @param int startRow 数据开始行
	 * @throws IOException 
	 * @throws InvalidFormatException 
	 * @throws EncryptedDocumentException ∂
	 * */
	public static List<HashMap<String, Object>> readExcelWithDoubleTitle(String filePath,int startRow) 
	        throws IOException, EncryptedDocumentException, InvalidFormatException {
		List<HashMap<String,Object>> list = new ArrayList<>();
		FileInputStream inputStream = new FileInputStream(filePath);
		 Workbook workBook=WorkbookFactory.create(inputStream);
//		HSSFWorkbook workBook = new HSSFWorkbook(inputStream);
        // 读取工作表
        Sheet sheet = workBook.getSheetAt(0);
        if (sheet.getLastRowNum() > 0) {
        	    Row title_row = sheet.getRow(startRow - 1);
        	    for (int t = startRow; t <= sheet.getLastRowNum() && sheet.getRow(t) != null; t++) {
        	    	 int cellNums = sheet.getRow(t).getLastCellNum();
                 Row row = sheet.getRow(t);
               
                 LinkedHashMap<String, Object> map = new LinkedHashMap<>();
                 for (int i = 0; i < cellNums; i++) {
                     if (null == row.getCell(i)) {
                         map.put(title_row.getCell(i).getStringCellValue(), "");
                         continue;
                     }
                     CellType cellType = row.getCell(i).getCellTypeEnum();
                     // NUMERIC格式的需要判断一下
                     if (cellType == CellType.NUMERIC) {
                         if(HSSFDateUtil.isCellDateFormatted(row.getCell(i))){
                             SimpleDateFormat df = new SimpleDateFormat("YYYY-MM-dd HH:mm:ss");
                             map.put(title_row.getCell(i).getStringCellValue(), df.format(row.getCell(i).getDateCellValue()));
                         } else {
                             String value = Helper.doubleToString(row.getCell(i).getNumericCellValue());
                             map.put(title_row.getCell(i).getStringCellValue(), value);
                         }
                     } else if (cellType == CellType.STRING)  {
                         map.put(title_row.getCell(i).getStringCellValue(), row.getCell(i).getStringCellValue());
                     } else if (cellType == CellType.FORMULA) {
                         String value = Helper.doubleToString(row.getCell(i).getNumericCellValue());
                         map.put(title_row.getCell(i).getStringCellValue(), value);
                     }
                 }
                 list.add(map);
        	    }
        	    inputStream.close();
            workBook.close();
        	    
        }
        return list;
	}
	
	
}
