package utils;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;

import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;

public class ExcelImportHSSFUtil {

	private HSSFFormulaEvaluator formulaEvaluator;
	private HSSFSheet sheet;
	
	
	public HSSFFormulaEvaluator getFormulaEvaluator() {
		return formulaEvaluator;
	}

	public void setFormulaEvaluator(HSSFFormulaEvaluator formulaEvaluator) {
		this.formulaEvaluator = formulaEvaluator;
	}

	public HSSFSheet getSheet() {
		return sheet;
	}

	public void setSheet(HSSFSheet sheet) {
		this.sheet = sheet;
	}

	public ExcelImportHSSFUtil() {
		super();
	}
	
	public ExcelImportHSSFUtil(InputStream is) throws IOException {
		this(is, 0, true);
	}
	
	public ExcelImportHSSFUtil(InputStream is, int seetIndex) throws IOException {
		this(is, seetIndex, true);
	}
	
	public ExcelImportHSSFUtil(InputStream is, int seetIndex, boolean evaluateFormular) throws IOException {
		super();
		HSSFWorkbook workbook = new HSSFWorkbook(is);
		this.sheet = workbook.getSheetAt(seetIndex);
		if(evaluateFormular){
			this.formulaEvaluator = new HSSFFormulaEvaluator(workbook);
		}
	}
	
	public Object getCellValue(Cell cell, int cellType) throws Exception{
		
		
		switch (cellType) {
		case Cell.CELL_TYPE_NUMERIC://0

			if(HSSFDateUtil.isCellDateFormatted(cell)){
				SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
				//return sdf.format(HSSFDateUtil.getJavaDate(cell.getNumericCellValue()));
				return sdf.format(cell.getDateCellValue());
			}else{
				//DecimalFormat df = new DecimalFormat("#");
				//return df.format(cell.getNumericCellValue());
				return cell.getNumericCellValue();
			}
		case Cell.CELL_TYPE_STRING://1
			return cell.getStringCellValue();
		case Cell.CELL_TYPE_FORMULA://2
			
			if(this.formulaEvaluator == null){//得到公式
				return cell.getCellFormula();
			}else{//计算公式
				CellValue evaluate = this.formulaEvaluator.evaluate(cell);
				cellType = evaluate.getCellType();
				return this.getCellValue(cell, cellType);
			}
		case Cell.CELL_TYPE_BLANK://3
			//注意空和没有值不一样，从来没有录入过内容的单元格不属于任何数据类型，不会走这个case
			return "";
		case Cell.CELL_TYPE_BOOLEAN://4
			return cell.getBooleanCellValue();
		case Cell.CELL_TYPE_ERROR:
			throw new Exception("数据类型错误");
		default:
			throw new Exception("未知数据类型:" + cellType);
		}
	}
}
