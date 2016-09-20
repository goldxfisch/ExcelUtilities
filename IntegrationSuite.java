/**
 * 
 */
package utilties;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * @author Vikas Bhatia
 *
 */
public class IntegrationSuite {

	/**
	 * @param args
	 */
	public static void main(String[] args) throws IOException {
	    String srcExcelFilePath = "C://OasisExecutions/ReUsable_working/Generic/Generic_working.xls";
	    String tgtExcelFilePath = "C://OasisExecutions/ReUsable/Generic/Generic.xls";
	    
	    readReusables reader = new readReusables();
	    List<ReusableBooks> srclistBooks = reader.readBooksFromExcelFile(srcExcelFilePath,tgtExcelFilePath );
	    System.out.println(srclistBooks);
	    
	//    List<ReusableBooks> tgtlistBooks = reader.readBooksFromExcelFile(tgtExcelFilePath);
	//    System.out.println(tgtlistBooks);
	}
	private Workbook getWorkbook(FileInputStream inputStream, String excelFilePath)
	        throws IOException {
	    Workbook workbook = null;
	 
	    if (excelFilePath.endsWith("xlsx")) {
	        workbook = new XSSFWorkbook(inputStream);
	    } else if (excelFilePath.endsWith("xls")) {
	        workbook = new HSSFWorkbook(inputStream);
	    } else {
	        throw new IllegalArgumentException("The specified file is not Excel file");
	    }
	 
	    return workbook;
	}
}
