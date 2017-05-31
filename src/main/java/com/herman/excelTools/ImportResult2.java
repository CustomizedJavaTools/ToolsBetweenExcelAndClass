package com.herman.excelTools;

import java.util.List;

public class ImportResult2<T> {

	private List<ImportResultEachRow<T>> resultList;
	
	private int errorStatus;

	public List<ImportResultEachRow<T>> getResultList() {
		return resultList;
	}

	public void setResultList(List<ImportResultEachRow<T>> resultList) {
		this.resultList = resultList;
	}
	
	public int getErrorStatus() {
		return errorStatus;
	}

	public void setErrorStatus(int errorStatus) {
		this.errorStatus = errorStatus;
	}
	
	public ImportResult2(List<ImportResultEachRow<T>> result, int errorStatus) {
		// TODO Auto-generated constructor stub
		this.resultList = result;
		this.errorStatus = errorStatus;
	}

}
