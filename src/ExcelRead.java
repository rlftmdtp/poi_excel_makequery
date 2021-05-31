import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelRead {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		
		FileInputStream fi = new FileInputStream("C:/Users/GILLSEUNGSAE/Desktop/test/test2.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fi);
		XSSFSheet sheet = workbook.getSheetAt(0);
		
		
		for(int i=1; i<sheet.getLastRowNum(); i++){
			XSSFRow row = sheet.getRow(i);
			if(row != null){
				List<String> cellList = new ArrayList<String>();
				
				// 35?
				StringBuffer sb = new StringBuffer();
				sb.append("INSERT INTO sn_gzt_mem_kipo(KIPO_RQSTR_ID, KIPO_RQSTR_NM, RQSTR_EMADR, EMAIL_LTST_TRSMS_DT, APAGT_CD, EMAIL_ROAMS_AGREE_YN, MOTN_DT) VALUES();");
				
				StringBuffer sb2 = new StringBuffer();
				for(int j=0; j<row.getLastCellNum(); j++){

					XSSFCell cell = row.getCell(j);
					CellType ct = cell.getCellTypeEnum();
					if(cell != null && (j==0 || j==1 || j==2 || j==5 || j==22 || j==25 || j==26)){
						if(ct != null){
							switch(cell.getCellTypeEnum()){
								case FORMULA:
									sb2.append("'"+cell.getCellFormula()+"',");
									break;
								case NUMERIC:
									cell.setCellType(Cell.CELL_TYPE_STRING); // 정수로 변환 
									sb2.append("'"+cell.getStringCellValue()+ ""+"',");
									break;
								case STRING:
									sb2.append("'"+cell.getStringCellValue()+ ""+"',");
									break;
								case BOOLEAN:
									sb2.append("'"+cell.getBooleanCellValue()+ ""+"',");
									break;
								case ERROR:
									sb2.append("'"+cell.getErrorCellValue()+ ""+"',");
									break;
							}
						}
					}
					// INSERT INTO sn_gzt_mem_kipo VALUES('KIPO20210001', '길승세', 'rlftmdtp@kipi.or.kr', '119981042713', '20210217', '20210226', '1');
				}

				StringBuffer lastDeleteSB = sb2.deleteCharAt(sb2.length()-1);
				sb.insert(140, lastDeleteSB.toString());
				System.out.println(sb.toString());
			}
		}
		
	}

	private static void swtich() {
		// TODO Auto-generated method stub
		
	}

}
