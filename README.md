# ToolsBetweenExcelAndClass
For converting excel to class or the other way around.

Example:

1. Class file

		//mapping between class and excel
		public class ExampleDTO{
			@ExcelColumn(column="A",maxLength=50,required=true,header="Test Column 1")
			public String TestCol1;
			@ExcelColumn(column="B",maxLength=100,required=true,header="Test Column 2")
			public String TestCol2;
		}

2.From excel to class:

	//copy to class list
	List<ImportResultEachRow<ExampleDTO>> result = new ArrayList<ImportResultEachRow<ExampleDTO>>();
	result=Excel2ClassUtil.copyExcel2Class(xssfWorkbook,ExampleDTO,2);

	//get target class list
	for(ImportResultEachRow<ExampleDTO> item : result){
		ExampleDTO exampleDTO=item.getTargetObject();
	}
    
3.From MultipartHttpServletRequest to class:

	ImportResult2<ExampleDTO> importResult = null;
	try {
		importResult = Excel2ClassUtil.importFile(model, request, ExampleDTO.class,2);
	} catch (Exception e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	}

	if (importResult.getErrorStatus()==1){
		system.out.println("File is not 'xlsx, xls' or file has error");
	}

	//save upload dto and check if there is invalid line
	List<String> invalidLines=new LinkedList<String>();
	for(ImportResultEachRow<ExampleDTO> item : importResult.getResultList()){
		ExampleDTO ExampleDTO=item.getTargetObject();
	}

4.From class to excel

	List<ExampleDTO> result=new ArrayList<ExampleDTO>();
	String tabName="BSRD Account";
	Workbook workbook=Class2ExcelUtil.class2Excel(ExampleDTO.class,result, false, tabName);
