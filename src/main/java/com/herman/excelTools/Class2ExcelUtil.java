package com.hsbc.gbm.ptutilities.common.util.excelTools;

import java.lang.reflect.Field;
import java.util.List;

import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Class2ExcelUtil.java
 * Purpose:	convert a class to a excel row.
 *
 * @author	Herman H WEN
 * @version 1.0 2017/05/17
 */

public class Class2ExcelUtil {

	public static <T> Workbook  class2Excel(Class<T> TargetClass,List<T> sources,boolean isGreaterThan07,String tabName){
		Workbook  xssfWorkbook = isGreaterThan07? new XSSFWorkbook():new HSSFWorkbook();
		Sheet sheet = xssfWorkbook.createSheet(tabName);
		Field[] fields=TargetClass.getDeclaredFields();

		//write header
		Row headerRow=sheet.createRow(0);
		for(Field field: fields){
			if (field.isAnnotationPresent(ExcelColumn.class)){
				ExcelColumn excelColumnAnnotation=field.getAnnotation(ExcelColumn.class);

				String header="";
				if(StringUtils.isEmpty(excelColumnAnnotation.header()))
					header=field.getName();
				else
					header=excelColumnAnnotation.header();

				Cell cell=headerRow.createCell(CommonTools.cStr2Num(excelColumnAnnotation.column())-1);
				cell.setCellValue(header);

				/*
				 * make the cell to be bold
				 */
				CellStyle style = xssfWorkbook.createCellStyle();//Create style
				Font font = xssfWorkbook.createFont();//Create font
				font.setBoldweight(Font.BOLDWEIGHT_BOLD);//Make font bold
				style.setFont(font);//set it to bold
				cell.setCellStyle(style);
			}
		}

		//write data
		for(int rowNum=1;rowNum<=sources.size();++rowNum){
			Row row=sheet.createRow(rowNum);
			//put data into fields
			for(Field field: fields){
				if (field.isAnnotationPresent(ExcelColumn.class)){
					ExcelColumn excelColumnAnnotation=field.getAnnotation(ExcelColumn.class);

					Cell cell=row.createCell(CommonTools.cStr2Num(excelColumnAnnotation.column())-1);
					try {
						boolean isAccessible=field.isAccessible();
						field.setAccessible(true);
						Object val=field.get(sources.get(rowNum-1));
						if(val==null) continue;
						String str=(String) val.toString();
						if(StringUtils.isNotEmpty(str)){
							cell.setCellValue(str);
						}
						field.setAccessible(isAccessible);


					} catch (IllegalArgumentException | IllegalAccessException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}

				}
			}
		}

		//auto fit each columns
		for(int col=0;col<CommonTools.getMaxColumn(TargetClass);++col)
			sheet.autoSizeColumn(col);


		return xssfWorkbook;

	}
}
