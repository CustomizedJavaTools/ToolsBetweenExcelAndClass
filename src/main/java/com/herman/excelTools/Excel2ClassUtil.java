package com.hsbc.gbm.ptutilities.common.util.excelTools;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.ui.Model;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.MultipartHttpServletRequest;

/**
 * Excel2ClassUtil.java
 * Purpose:	convert excel row to a class.
 *
 * @author	Herman H WEN
 * @version 1.0 2017/05/17
 */

public class Excel2ClassUtil {

	public static <T> ImportResult2<T> importFile(Model model, MultipartHttpServletRequest request,Class<T> TargetClass,int startRow) throws Exception {

		//check if max column reach limit
		if (CommonTools.getMaxColumn(TargetClass)>16384) throw new Exception("Column reach limit!");

		List<ImportResultEachRow<T>> result = new ArrayList<ImportResultEachRow<T>>();

		//get workbook
		Workbook xssfWorkbook;

		//copy Excel to Class
		try {
			xssfWorkbook = extractWorkbookFrom(request);
			if (xssfWorkbook==null) return new ImportResult2<T>(result,1);
			result=copyExcel2Class(xssfWorkbook,TargetClass,startRow);
		} catch (IOException | InstantiationException | IllegalAccessException e) {
			return new ImportResult2<T>(result,1);
		}

		return new ImportResult2<T>(result,0);
	}

	//extract workbook from MultipartHttpServletRequest
	public static Workbook extractWorkbookFrom(MultipartHttpServletRequest request) throws IOException{
		InputStream is = null;
		MultipartFile file = request.getFile("file");  
		String fileName = file.getOriginalFilename();
		String[] fileNameSplits = fileName.split("\\.");
		String sub = "";

		if (fileNameSplits.length>0) {
			sub = fileNameSplits[fileNameSplits.length-1];
		}
		if(sub.equalsIgnoreCase("xlsx")||sub.equalsIgnoreCase("xls")){
			is = file.getInputStream();
			Workbook  xssfWorkbook = null;
			return sub.equalsIgnoreCase("xlsx")? new XSSFWorkbook(is):new HSSFWorkbook(is);
		}
		return null;
	}

	//check if all header present in worksheet row 1
	public static List<String> checkHeader(Workbook  xssfWorkbook,Class TargetClass){
		List<String> missingHeaders=new ArrayList<String>();
		for(int index = 0; index < xssfWorkbook.getNumberOfSheets(); index++) {
			Sheet xssfSheet = xssfWorkbook.getSheetAt(index);
			if(xssfSheet == null) continue;
			Row xssfRow = xssfSheet.getRow(0);
			Field[] fields=TargetClass.getDeclaredFields();

			//check if all header present in worksheet row 1
			for(Field field: fields){
				if (field.isAnnotationPresent(ExcelColumn.class)){
					ExcelColumn excelColumnAnnotation=field.getAnnotation(ExcelColumn.class);
					if(StringUtils.isEmpty(excelColumnAnnotation.header())) continue;
					if (!xssfRow.getCell(CommonTools.cStr2Num(excelColumnAnnotation.column())).equals(excelColumnAnnotation.header()))
						missingHeaders.add(excelColumnAnnotation.header());
				}
			}
		}
		return missingHeaders;
	}

	//copy all worksheets from excel to a list of class
	public static <T> List<ImportResultEachRow<T>>  copyExcel2Class(Workbook  xssfWorkbook,Class<T> TargetClass,int startRow) throws InstantiationException, IllegalAccessException{
		List<ImportResultEachRow<T>> result = new ArrayList<ImportResultEachRow<T>>();
		for(int index = 0; index < xssfWorkbook.getNumberOfSheets(); index++) {
			Sheet xssfSheet = xssfWorkbook.getSheetAt(index);
			if(xssfSheet == null) continue;
			for (int rowNum = startRow-1; rowNum <= xssfSheet.getLastRowNum(); rowNum++) {
				
				Row xssfRow = xssfSheet.getRow(rowNum);
				
				//stop when meeting empty row.
				if (checkIfEmptyRow(CommonTools.getMaxColumn(TargetClass),xssfRow)) 
					break;
				
				ImportResultEachRow resultEachRow;
				resultEachRow = row2Object(TargetClass,xssfRow);
				if (resultEachRow!=null){
					result.add(resultEachRow);
				}
			}
		}
		return result;
	}

	//copy row to a class
	private static <T> ImportResultEachRow row2Object(Class<T> TargetClass,Row xssfRow) throws InstantiationException, IllegalAccessException{

		ImportResultEachRow result=new ImportResultEachRow();
		List<String> overMaxLengthColumns= new ArrayList<String>();
		List<String> overMaxLengthColumnsWithName= new ArrayList<String>();
		List<String> notReachMinLengthColumns= new ArrayList<String>();
		List<String> notReachMinLengthColumnsWithName= new ArrayList<String>();
		List<String> emptyColumns = new ArrayList<String>();
		List<String> emptyColumnsWithName = new ArrayList<String>();
		T target = TargetClass.newInstance();

		//get max column
		int maxCoIx=CommonTools.getMaxColumn(TargetClass);

		//check if it is an empty row
		if (checkIfEmptyRow(maxCoIx,xssfRow)) return null;

		//first column
		int minCoIx = xssfRow.getFirstCellNum();
		Field[] fields=TargetClass.getDeclaredFields();

		//put data into fields
		for(Field field: fields){
			if (field.isAnnotationPresent(ExcelColumn.class)){
				ExcelColumn excelColumnAnnotation=field.getAnnotation(ExcelColumn.class);
				
				//get header
				String header=null;
				if(StringUtils.isEmpty(excelColumnAnnotation.header()))
					header=field.getName();
				else
					header=excelColumnAnnotation.header();
				
				boolean gotValue = false;
				Cell cell = xssfRow.getCell(CommonTools.cStr2Num(excelColumnAnnotation.column())-1);
				if(cell !=null ) {

					String val="";
					if(cell.getCellType() == Cell.CELL_TYPE_NUMERIC){
						DecimalFormat df = new DecimalFormat(); 
						val = df.format(cell.getNumericCellValue());
						val = val.replace(",", "");
					}else {
						val=cell.toString();
					}
			
					//cut the tail if value is too long
					if(excelColumnAnnotation.maxLength()!=-1 && val.length()>excelColumnAnnotation.maxLength()){
						val=val.substring(0, excelColumnAnnotation.maxLength());
						overMaxLengthColumns.add(excelColumnAnnotation.column());	
						overMaxLengthColumnsWithName.add(header);
					}
					
					/*
					 * check if larger than or equal to min length
					 * if max length is smaller than min length then ignore min length
					 */
					if(val.length()>=excelColumnAnnotation.minLength() || excelColumnAnnotation.maxLength()<excelColumnAnnotation.minLength()){
						if(!(excelColumnAnnotation.required() && org.springframework.util.StringUtils.isEmpty(val))) {
							//put value into field of class
							boolean isAccessible=field.isAccessible();
							field.setAccessible(true);
							if(field.getType().isAssignableFrom(String.class)){//if the type of field is String then set the value is val
								field.set(target, val);
							}else{
								/*
								 * if the type of field is not String , try to find its constructor with String parameter.
								 * if found , put the val into field through its constructor.
								 */
								try {
									Class targetFieldClass=field.getType();
									Constructor constructor= targetFieldClass.getConstructor(String.class);
									if (constructor!=null){
										try {
											Object targetField=constructor.newInstance(val);
											field.set(target, targetField);
										} catch (IllegalArgumentException
												| InvocationTargetException e) {
											// TODO Auto-generated catch block
											e.printStackTrace();
										}
									}
								} catch (NoSuchMethodException | SecurityException e) {
									// TODO Auto-generated catch block
									e.printStackTrace();
								}
								//TODO add another type
							}
							
							field.setAccessible(isAccessible);
							gotValue=true;
						}
					}else{
						/*
						 * if the length of value is shorter than min length
						 */
						notReachMinLengthColumns.add(excelColumnAnnotation.column());	
						notReachMinLengthColumnsWithName.add(header);
					}
				}
				if (excelColumnAnnotation.required() && !gotValue){
					emptyColumnsWithName.add(header);					
					emptyColumns.add(excelColumnAnnotation.column());
				}
			}
		}
		
		//put needed result into reult object
		result.setLine(xssfRow.getRowNum()+1);
		result.setTargetObject(target);
		
		result.setEmptyColumnsWithName(emptyColumnsWithName);
		result.setEmptyColumns(emptyColumns);
		
		result.setOverMaxLengthColumns(overMaxLengthColumns);
		result.setOverMaxLengthColumnsWithName(overMaxLengthColumnsWithName);
		
		result.setNotReachMinLengthColumns(notReachMinLengthColumns);
		result.setNotReachMinLengthColumnsWithName(notReachMinLengthColumnsWithName);
		
		if(emptyColumns.size()>0) result.setMissing(true);
		if(overMaxLengthColumns.size()>0) result.setOverMaxLength(true);
		if(notReachMinLengthColumns.size()>0) result.setNotReachMinLength(true);
		
		return result;

	}

	//check if it is an empty row
	private static boolean checkIfEmptyRow(int maxCoIx,Row xssfRow){
		if (xssfRow==null) return true;
		int minCoIx=xssfRow.getFirstCellNum();
		for (int colIx = minCoIx; colIx <= maxCoIx; colIx++) {
			Cell cell = xssfRow.getCell(colIx);
			if(cell!=null)
				if (StringUtils.isNotEmpty(cell.toString().trim()))
					return false;
		}
		return true;
	}


}
