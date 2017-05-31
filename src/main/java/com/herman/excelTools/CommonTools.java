package com.hsbc.gbm.ptutilities.common.util.excelTools;

import java.lang.reflect.Field;
import java.util.LinkedList;
import java.util.List;

import org.apache.commons.lang.StringUtils;
/**
 * CommonTools.java
 * Purpose:	some common tools for the convenience of converting between excel to class.
 *
 * @author	Herman H WEN
 * @version 1.0 2017/05/17
 */
public class CommonTools {

	//get max column's number from annotation
		static int getMaxColumn(Class TargetClass){
			Field[] fields=TargetClass.getDeclaredFields();

			//get max column
			int maxCoIx=0;
			for(Field field: fields){
				if (field.isAnnotationPresent(ExcelColumn.class)){
					ExcelColumn excelColumnAnnotation=field.getAnnotation(ExcelColumn.class);
					int col=cStr2Num(excelColumnAnnotation.column());
					if(cStr2Num(excelColumnAnnotation.column())>maxCoIx)
						maxCoIx=col;
				}
			}
			return maxCoIx;
		}

		/*
		 * Convert string to a num
		 * e.g. A=1,B=2...,AA=27...,AAA=703
		 * ABC=1*26^2+2*26^1+3^260=731
		 */
		static int cStr2Num(String column){
			int sum=0;
			char[] chars=column.toUpperCase().toCharArray();
			for (int i = 0; i < chars.length; i++)
				sum+=char2Num(chars[i])*Math.pow(26, chars.length-i-1);
			return sum;
		}

		/*
		 * convert char to a num
		 * A=1,B=2,C=3...,Z=26
		 */
		private static int char2Num(char chr){
			return chr-64;
		}
		
		
		public static String joinList(List<String> list,String separator){
			List<String> tmpList=new LinkedList<String>();
			if(list.size()<4){
				return StringUtils.join(list,separator);
			}else{
				tmpList.add(list.get(0));
				tmpList.add(list.get(1));
				tmpList.add("...");
				tmpList.add(list.get(list.size()-1));
				return StringUtils.join(tmpList,separator);
			}
		}
}
