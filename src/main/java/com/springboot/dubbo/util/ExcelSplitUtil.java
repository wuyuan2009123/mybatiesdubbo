package com.springboot.dubbo.util;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.impl.piccolo.io.FileFormatException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;


/**
 * 支持 2003 和 2007 分割 
 * @author wuy
 *
 */
public class ExcelSplitUtil {
	
	public static final String EXTENSION_XLS = "xls";
	public static final String EXTENSION_XLSX = "xlsx";
	
	private static Logger LOG = LoggerFactory.getLogger(ExcelSplitUtil.class);
	
	public static void main(String[] args) throws Exception {
		String filePath="C:\\Users\\wuy\\Desktop\\instrint_history\\调查单号_2017_11_11.xls";
		System.out.println(filePath.substring(filePath.lastIndexOf(".")));
		int splitRows=50;
		split(filePath,splitRows);
    }


	private static void split(String filePath,int splitRows) throws IOException {
		String outPath=filePath.substring(0,filePath.lastIndexOf("\\")+1);
		String extPath=filePath.substring(filePath.lastIndexOf("."));
		Workbook workbook = getWorkbook(filePath);
		for (int numSheet = 0; numSheet < workbook.getNumberOfSheets(); numSheet++) {
			org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheetAt(numSheet);
			if (sheet == null) {
				continue;
			}
			int firstRowIndex = sheet.getFirstRowNum();
			int lastRowIndex = sheet.getLastRowNum();
			List<Map<String, String>> excelList=null;
			int totalPage=(lastRowIndex%splitRows)==0?(lastRowIndex/splitRows):  (lastRowIndex/splitRows)+1;
			System.out.println("开始读取 Excel 页签" + numSheet);
			int startRow=0;
			int lastPageRow=0;
			int pageCount=0;
			for (int rowIndex = firstRowIndex + 1; rowIndex <= lastRowIndex; rowIndex++) {
				Row currentRow = sheet.getRow(rowIndex);
				Map<String, String> map=new TreeMap<String, String>();
				
				String userId = getCellValue(currentRow.getCell(1), true);
				map.put("tradeId", userId);
				
				if(pageCount<= totalPage){
					if(pageCount<=totalPage-1){
						startRow++;
						if(startRow==1){
							pageCount++;
							excelList=new ArrayList<>();
							excelList.add(map);
						}else if(startRow==splitRows){
							startRow=0;
							excelList.add(map);
							if(pageCount==totalPage-1){
								pageCount++;
								writeXlsOrXlsxExcelFile(outPath+(pageCount-1)+extPath,new String[] { "tradeId"}, new String[]{"tradeId"}, excelList);
							}else{
								writeXlsOrXlsxExcelFile(outPath+pageCount+extPath,new String[] { "tradeId"}, new String[]{"tradeId"}, excelList);
							}
						}
						else{
							excelList.add(map);
						}
					}
					else if(pageCount==totalPage){
						lastPageRow++;
						if(lastPageRow==1){
							excelList=new ArrayList<>();
						}
						excelList.add(map);
					}
				}
			}
			//最后一个页签
			writeXlsOrXlsxExcelFile(outPath+pageCount+extPath,new String[] { "tradeId"}, new String[]{"tradeId"}, excelList);
			
		}
	}

	
	/***
     * <pre>
     * 取得Workbook对象(xls和xlsx对象不同,不过都是Workbook的实现类)
     *   xls:HSSFWorkbook
     *   xlsx：XSSFWorkbook
     * @param filePath
     * @return
     * @throws IOException
     * </pre>
     */
	public static Workbook getWorkbook(String filePath) throws IOException {
        Workbook workbook = null;
        InputStream is = new FileInputStream(filePath);
        if (filePath.endsWith(EXTENSION_XLS)) {
            workbook = new HSSFWorkbook(is);
        } else if (filePath.endsWith(EXTENSION_XLSX)) {
            workbook = new XSSFWorkbook(is);
        }
        return workbook;
    }
	
	/**
     * 文件检查
     * @param filePath
     * @throws FileNotFoundException
     * @throws FileFormatException
     */
    public void preReadCheck(String filePath) throws FileNotFoundException, FileFormatException {
        // 常规检查
        File file = new File(filePath);
        if (!file.exists()) {
            throw new FileNotFoundException("传入的文件不存在：" + filePath);
        }

        if (!(filePath.endsWith(EXTENSION_XLS) || filePath.endsWith(EXTENSION_XLSX))) {
            throw new FileFormatException("传入的文件不是excel");
        }
    }
	
    
    /**
     * 取单元格的值
     * @param cell 单元格对象
     * @param treatAsStr 为true时，当做文本来取值 (取到的是文本，不会把“1”取成“1.0”)
     * @return
     */
    public static String getCellValue(Cell cell, boolean treatAsStr) {
        if (cell == null) {
            return "";
        }
        
        if (treatAsStr) {
            // 虽然excel中设置的都是文本，但是数字文本还被读错，如“1”取成“1.0”
            // 加上下面这句，临时把它当做文本来读取
            cell.setCellType(Cell.CELL_TYPE_STRING);
        }

        if (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
            return String.valueOf(cell.getBooleanCellValue());
        } else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
            return String.valueOf(cell.getNumericCellValue());
        } else {
            return String.valueOf(cell.getStringCellValue());
        }
    }
    
    
    public static void setCellValue(Cell newCell, Cell cell, HSSFWorkbook wb) {
        if (cell == null) {
            return;
        }
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_BOOLEAN:
                newCell.setCellValue(cell.getBooleanCellValue());
                break;
            case Cell.CELL_TYPE_NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    HSSFCellStyle cellStyle = wb.createCellStyle();
                    HSSFDataFormat format = wb.createDataFormat();
                    cellStyle.setDataFormat(format.getFormat("yyyy/mm/dd HH:mm:ss"));
                    newCell.setCellStyle(cellStyle);
                    newCell.setCellValue(cell.getDateCellValue());
                } else {
                    newCell.setCellValue(cell.getNumericCellValue());
                }
                break;
            case Cell.CELL_TYPE_FORMULA:
                newCell.setCellValue(cell.getCellFormula());
                break;
            case Cell.CELL_TYPE_STRING:
                newCell.setCellValue(cell.getStringCellValue());
                break;
        }
    }
    
    /**
	 * 设置样式
	 * @param workbook
	 * @return
	 */
    public static CellStyle getStyle(Workbook workbook) {
		CellStyle style = workbook.createCellStyle();
		style.setAlignment(CellStyle.ALIGN_CENTER);
		style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		// 设置单元格字体
		Font headerFont = workbook.createFont(); // 字体
		headerFont.setFontHeightInPoints((short) 14);
		headerFont.setColor(HSSFColor.RED.index);
		headerFont.setFontName("宋体");
		style.setFont(headerFont);
		style.setWrapText(true);

		// 设置单元格边框及颜色
		style.setBorderBottom((short) 1);
		style.setBorderLeft((short) 1);
		style.setBorderRight((short) 1);
		style.setBorderTop((short) 1);
		style.setWrapText(true);
		return style;
	}
	
	public static void writeXlsOrXlsxExcelFile(String filePath,String[] titleArray,String[] fieldArray,
			List<Map<String, String>> excelList){
		try {
			Workbook workbook = null;
			if (filePath.endsWith(EXTENSION_XLS)) {
	            workbook = new HSSFWorkbook();
	        } else if (filePath.endsWith(EXTENSION_XLSX)) {
	            workbook = new XSSFWorkbook();
	        }
			Sheet sheet = workbook.createSheet("数据");
			// 创建第一栏
			Row headRow = sheet.createRow(0);
			// 获取参数个数作为excel列数
			int columeCount = titleArray.length;
			for (int m = 0; m < columeCount; m++) {
				Cell cell = headRow.createCell(m, Cell.CELL_TYPE_STRING);
				CellStyle style = getStyle(workbook);
				cell.setCellStyle(style);
				cell.setCellValue(titleArray[m]);
				sheet.autoSizeColumn(m);
			}
			for (int i = 0; i < excelList.size(); i++) {
				Map<String, String> map = excelList.get(i);
				// 写入数据
				Row row = sheet.createRow(i + 1);
				for (int k = 0; k < titleArray.length; k++) {
					row.createCell(k, Cell.CELL_TYPE_STRING);
					row.getCell(k).setCellValue(trimToEmpty(map.get(fieldArray[k])).replace("\n", ""));
				}
			}
			FileOutputStream fileOutputStream = new FileOutputStream(new File(filePath));
			workbook.write(fileOutputStream);
			fileOutputStream.close();
		} catch (Exception e) {
			LOG.error("XlsOrXlsx 文件写入异常：" + e);
			LOG.error(e.getMessage(),e);
        }
	}
	
	public static String trimToEmpty(Object ob) {
        return ob == null ? "" : ob.toString().trim();
    }
    
    /**
     * 写csv文件
     */
	/*
    public static void storeIntoCsv(String fileName,List<String> dataList){
    	Writer writer = null;
        final String NEW_LINE = "\n";
        try {     
            StringBuilder csvStr = new StringBuilder();             
            //数据行
            for(String csvData : dataList){
                csvStr.append(csvData).append(NEW_LINE);
            }           
            //写文件
            writer = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(new File(fileName)), "GB2312"));
            writer.write(csvStr.toString());
            writer.flush();
            writer.close();
        } catch (Exception e) {
            log.error(e.getMessage(), e);
        } finally {
        	IOUtils.closeQuietly(writer);
        }       
    }*/
    
    /**
	 * 写csv文件
	 * @param fileName
	 * @param lineDataList
	 */
	public static void writeCSVExcelFile(String fileName, List<List<String>> lineDataList) {
        final String new_line = "\n";
        try {
            File file = new File(fileName);
            // 如果不存在,创建一个新文件
            if (!file.exists()) {
                file.createNewFile();
            }
            StringBuilder csvStr = new StringBuilder();
            // 数据行
            for (List<String> lineData : lineDataList) {
                String csvData = "";
                for (String metaData : lineData) {
                    if (metaData.indexOf(',') > 0) {
                        metaData = metaData.replace(',', '.');
                    }
                    if (!csvData.endsWith(",") && !"".equals(csvData)) {
                        csvData += ",";
                    }
                    csvData += metaData;
                }
                csvStr.append(csvData).append(new_line);
                lineData.clear();
            }
            // 写文件
            Writer writer = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(
                    new File(fileName), true), "GB2312"));
            writer.write(csvStr.toString());
            writer.flush();
            writer.close();

            lineDataList.clear();
        } catch (Exception e) {
        	LOG.error("CSV 文件写入异常：" + e);
        	LOG.error(e.getMessage(),e);
        }
    }

	
	
}
