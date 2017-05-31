package com.herman.excelTools;

import java.util.List;

public class ImportResultEachRow<T> {
	private T targetObject;
	private int line;
	private boolean isMissing;
	private List<String> emptyColumns;
	private List<String> emptyColumnsWithName;
	
	private boolean isOverMaxLength;
	private List<String> overMaxLengthColumns;
	private List<String> overMaxLengthColumnsWithName;
	
	private boolean isNotReachMinLength;
	private List<String> notReachMinLengthColumns;
	private List<String> notReachMinLengthColumnsWithName;
	public T getTargetObject() {
		return targetObject;
	}
	public void setTargetObject(T target) {
		targetObject = target;
	}
	public List<String> getEmptyColumns() {
		return emptyColumns;
	}
	public void setEmptyColumns(List<String> emptyColumns) {
		this.emptyColumns = emptyColumns;
	}
	public List<String> getEmptyColumnsWithName() {
		return emptyColumnsWithName;
	}
	public void setEmptyColumnsWithName(List<String> emptyColumnsWithName) {
		this.emptyColumnsWithName = emptyColumnsWithName;
	}
	public int getLine() {
		return line;
	}
	public void setLine(int line) {
		this.line = line;
	}
	public boolean isMissing() {
		return isMissing;
	}
	public void setMissing(boolean isMissing) {
		this.isMissing = isMissing;
	}
	public boolean isOverMaxLength() {
		return isOverMaxLength;
	}
	public void setOverMaxLength(boolean isOverMaxLength) {
		this.isOverMaxLength = isOverMaxLength;
	}
	public List<String> getOverMaxLengthColumns() {
		return overMaxLengthColumns;
	}
	public void setOverMaxLengthColumns(List<String> overMaxLengthColumns) {
		this.overMaxLengthColumns = overMaxLengthColumns;
	}
	public List<String> getOverMaxLengthColumnsWithName() {
		return overMaxLengthColumnsWithName;
	}
	public void setOverMaxLengthColumnsWithName(
			List<String> overMaxLengthColumnsWithName) {
		this.overMaxLengthColumnsWithName = overMaxLengthColumnsWithName;
	}
	public boolean isNotReachMinLength() {
		return isNotReachMinLength;
	}
	public void setNotReachMinLength(boolean isNotReachMinLength) {
		this.isNotReachMinLength = isNotReachMinLength;
	}
	public List<String> getNotReachMinLengthColumns() {
		return notReachMinLengthColumns;
	}
	public void setNotReachMinLengthColumns(List<String> notReachMinLengthColumns) {
		this.notReachMinLengthColumns = notReachMinLengthColumns;
	}
	public List<String> getNotReachMinLengthColumnsWithName() {
		return notReachMinLengthColumnsWithName;
	}
	public void setNotReachMinLengthColumnsWithName(
			List<String> notReachMinLengthColumnsWithName) {
		this.notReachMinLengthColumnsWithName = notReachMinLengthColumnsWithName;
	}
	
	
}
