package utilities;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.swing.text.html.HTMLDocument.Iterator;

import org.apache.poi.hssf.util.HSSFColor.HSSFColorPredefined;
import org.apache.poi.ss.usermodel.*;


public class ExcelUtils {
	private static Workbook wb;
	private static Sheet ws;
	private static Cell cell;
	private static Row row;
	private static CellStyle cellStyle;
	private static Color myColor;
	private static String excelFilePath;
	private static FileInputStream fileInput;
	private static FileOutputStream fileOut;

//	private Map<String, Integer> columns = new HashMap<String, Integer>();
//Do su dung thuong xuyen nen de public static
//	public ExcelUtils(String path) {
//		excelFilePath = path;
//	}

	//1. WORKING WITH EXCEL FILE

	public static void setExcelFile(String excelPath, String SheetName){
		try {
//			File f = new File(excelPath);
//
//			if (!f.exists()) {
//				f.createNewFile();
//				System.out.println("File doesn't exist, so created!");
//			}

			FileInputStream excelFile = new FileInputStream(excelPath);
			wb = WorkbookFactory.create(excelFile);
			ws = wb.getSheet(SheetName);
//			if (ws == null) {
//				ws = wb.createSheet(SheetName);
//			}
			excelFilePath = excelPath;

////			//Thêm tiêu đề tất cả các cột vào 'columns' map 
////			sh.getRow(i:0).forEach(cell ->{
////				columns.put(cell.getStringCellValue(), get.getColumnIndex());
////			});
//		}
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
	}

	public static Sheet getSheet(String SheetName) {
		 Sheet ws=wb.getSheet(SheetName);
		if(!isSheet(SheetName)) {
			ws=wb.createSheet(SheetName);
		}
		return ws;
	}

	public static void open() throws IOException{
		File file=new File(excelFilePath);
		if(file.canRead()) {
			FileInputStream fileInput=new FileInputStream(file);
			wb = WorkbookFactory.create(fileInput);
			fileInput.close();
		}
	}
	
	public static void save()throws IOException{
		FileOutputStream streamOut=new FileOutputStream(excelFilePath);
		wb.write(streamOut);
		streamOut.flush();
		streamOut.close();
	}

	public static void saveAs(String path)throws IOException{
		FileOutputStream streamOut=new FileOutputStream(path);
		wb.write(streamOut);
		streamOut.flush();
		streamOut.close();
	}
	
	public static boolean isSheet(String SheetName) {
		return wb.getSheetIndex(SheetName)>=0;
	}

	public static void addSheet(String SheetName) {
		if(!isSheet(SheetName)) {
			wb.createSheet(SheetName);
		}
	}
	
	public static void removeSheet(int SheetIndex) {
		wb.removeSheetAt(SheetIndex);
	}
	
	public static void removeSheet(String SheetName) {
		int index=wb.getSheetIndex(SheetName);
		removeSheet(index);
	}

	//2. WORKING WITH COLUMN IN EXCEL
	public static void addColumn(Sheet sheet, String ColName) {
		cellStyle=wb.createCellStyle();
		cellStyle.setFillForegroundColor(HSSFColorPredefined.RED.getIndex());
		cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		Row row=sheet.getRow(0);
		if(row==null)
			row=sheet.createRow(0);
		
		Cell cell;
		if(row.getLastCellNum()==-1) 
			cell=row.createCell(0);
		else
			cell=row.createCell(row.getLastCellNum());
		cell.setCellValue(ColName);
		//cell.setCellStyle(style);
	}

	public static void addColumn(String SheetName,String ColName) {
		Sheet ws=getSheet(SheetName);
		addColumn(ws, ColName);
	}
	
	public static void removeColumn(String SheetName, int colNum) {
		Sheet sheet=getSheet(SheetName);

		int rowCount=sheet.getLastRowNum()+1; 
		 
		Row row;
		for(int i=0;i<rowCount;i++) {
			row=sheet.getRow(i);
			if(row!=null) {
				Cell cell =row.getCell(colNum);
				if(cell!=null) {
					row.removeCell(cell);
				}
			}
		}
	}

	public static int convertColNameToColNum(Sheet sheet, String colName) {
		Row row=sheet.getRow(0);
		int cellRowNumber=row.getLastCellNum();
		int colNum=-1;
		
		for(int i=0;i<cellRowNumber;i++) {
			if(row.getCell(i).getStringCellValue().trim().equals(colName)) {
				colNum=i;
			}
		}
		return colNum;
	}

	//3. WORKING WITH ROW IN EXCEL
		public static Row getRow(Sheet sheet, int rowIndex) {
			Row row=sheet.getRow(rowIndex);
			if(row==null) {
				row=sheet.createRow(rowIndex);
			}
			return row;
		}
		
		// This method is to get the row count used of the excel sheet
		public static int getRowCount(String SheetName) {
			ws = wb.getSheet(SheetName);
			int number = ws.getLastRowNum() + 1;
			return number;
		}

		// This method is to get the Row number of the test case
		// This methods takes three arguments(Test Case name , Column Number & Sheet
		// name)
//		public static int getRowContains(String SheetName,String sTestCaseName, int colNum ) throws Exception {
//			int i;
//			ws = wb.getSheet(SheetName);
//			int rowCount = ExcelUtils.getRowCount(SheetName);
//			for (i = 0; i < rowCount; i++) {
//				if (getCellData(i, colNum).equalsIgnoreCase(sTestCaseName)) {
//					break;
//				}
//			}
//			return i;
//		}

		// This method is to get the count of the test steps of test case
		// This method takes three arguments (Sheet name, Test Case Id & Test case
		// row number)
//		public static int getTestStepsCount(String SheetName,String sTestCaseID, int iTestCaseStart) throws Exception {
//			for (int i = iTestCaseStart; i <= ExcelUtils.getRowCount(SheetName); i++) {
//				if (!sTestCaseID.equals(ExcelUtils.getCellValue(SheetName,i, Constants.Col_TestCaseID))) {
//					int number = i;
//					return number;
//				}
//			}
//			workSheet = workBook.getSheet(SheetName);
//			int number = workSheet.getLastRowNum() + 1;
//			return number;
//		}

		//4. WORKING WITH CELL IN EXCEL
		   //Chỉ số hàng và cột trên excel được tính từ 0
		public static String getCellData(int rownum, int colnum) {
			try {
//				FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
//				DataFormatter df = new DataFormatter();
				
				cell = ws.getRow(rownum).getCell(colnum);
				String cellData = null;
						
					switch (cell.getCellType()) {
						case STRING:
							cellData = cell.getStringCellValue();
							break;
						case NUMERIC:
							cellData = String.valueOf(cell.getNumericCellValue());
//							cellData= df.formatCellValue(cell);
							break;
						case BOOLEAN:
							cellData=Boolean.toString(cell.getBooleanCellValue());
							break;
						case BLANK:
							cellData="";
							break;
						case FORMULA:
							cellData=cell.getCellFormula();
							break;
						default:
							break;
					}
					return cellData;

			} catch (Exception e) {
				System.out.println(e.getMessage());
				return "";
			}
		}
		
		public static String getCellData(String sheetName, int rowIndex, int colIndex) {
			Sheet sheet = getSheet(sheetName);
			Row row=getRow(sheet, rowIndex);
			Cell cell=getCell(row, colIndex);
			return cell.getStringCellValue();
		} 
		
		public static Cell getCell(Row row, int colIndex) {
			Cell cell=row.getCell(colIndex-1);
			if(cell==null) {
				cell=row.createCell(colIndex-1);
			}
			return cell;
		}
		
		public static void setCell(Sheet sheet, int rowIndex, int colIndex, String value) {
			Row row=getRow(sheet, rowIndex);
			Cell cell=getCell(row,colIndex);
			cell.setCellValue(value);
		} 
		public static void setCell(String sheetName,int rowIndex,int colIndex, String value) {
			Sheet sheet=getSheet(sheetName);
			setCell(sheet, rowIndex, colIndex, value);
		}
		
		public static void setCell(String sheetName,String colName,int rowIndex, String value) {
			Sheet sheet=getSheet(sheetName);
			int colIndex=convertColNameToColNum(sheet,colName);
			setCell(sheet, rowIndex, colIndex, value);
		}
		
}
