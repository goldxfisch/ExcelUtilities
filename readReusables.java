package utilties;


import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class readReusables {


	 
	    private Object getCellValue(Cell cell) {
	        switch (cell.getCellType()) {
	        case Cell.CELL_TYPE_STRING:
	            return cell.getStringCellValue();
	     
	        case Cell.CELL_TYPE_BOOLEAN:
	            return cell.getBooleanCellValue();
	     
	        case Cell.CELL_TYPE_NUMERIC:
	        	 double d = cell.getNumericCellValue();
	        	 if (HSSFDateUtil.isCellDateFormatted(cell)) 
	        	 {
	        		 
	        		 Date javaDate= DateUtil.getJavaDate((double) d);
	        	        System.out.println(new SimpleDateFormat("MM/dd/yyyy").format(javaDate));
	        	        return(new SimpleDateFormat("MM/dd/yyyy").format(javaDate));
	        	
	               }
	        	 return String.valueOf(cell.getNumericCellValue()); 
	        	 
	        
	            
	       
	        }
	     
	        return null;
	    }
	    public List<ReusableBooks> readBooksFromExcelFile(String sourceExcelFilePath, String targetExcelFilePath) throws IOException {
	        List<ReusableBooks> listBooks = new ArrayList<>();
	        FileInputStream sourceInputStream = new FileInputStream(new File(sourceExcelFilePath));
	        FileInputStream targetInputStream = new FileInputStream(new File(targetExcelFilePath));
	    ///    Workbook workbook = new XSSFWorkbook(inputStream);
	        Workbook sourceWorkbook = new HSSFWorkbook(sourceInputStream);
	        Workbook targetWorkbook = new HSSFWorkbook(targetInputStream);
	        int numberOfSheetsSource = sourceWorkbook.getNumberOfSheets();
	        int numberOfSheetsTarget = sourceWorkbook.getNumberOfSheets();
	        Sheet tgtSheet;
	        Sheet srcSheet ;
	        System.out.println("CHECK 1: Number of Sheets in Souce <"+numberOfSheetsSource+">  Number of Sheets in Target<"+numberOfSheetsTarget+">");
	        for (int ctrSheet1 = 0; ctrSheet1< numberOfSheetsSource; ctrSheet1++) 
	        {
	             srcSheet = sourceWorkbook.getSheetAt(ctrSheet1);
	            for (int ctrSheet2 = 0; ctrSheet2< numberOfSheetsTarget; ctrSheet2++) 
	            {
	            	 tgtSheet = targetWorkbook.getSheetAt(ctrSheet2);
	            	if ( srcSheet.getSheetName().equals(tgtSheet.getSheetName()) )
	            	{
	            		 System.out.println("CHECK 2: MATCHING SOURCE WORKBOOK SHEET Number <"+ctrSheet1+">"+" SHEET Name <"+srcSheet.getSheetName()+">");
	         	        break;
	            	}
	            }
	        }
	         srcSheet = sourceWorkbook.getSheetAt(0);
	         tgtSheet = targetWorkbook.getSheetAt(0);
	         
	        Iterator<Row> iteratorSrcSheet = srcSheet.iterator();
	        Iterator<Row> iteratorTgtSheet = tgtSheet.iterator();
	        
	        int numRowsSrcWorkbook=0;
	        int numRowsTgtWorkBook=0;
	        while (iteratorSrcSheet.hasNext()) { numRowsSrcWorkbook++; iteratorSrcSheet.next();};
	        while (iteratorTgtSheet.hasNext()) { numRowsTgtWorkBook++;iteratorTgtSheet.next(); };
	        System.out.println("CHECK 3: MATCHING NUMBER OF ROWS :  SOURCE WORKBOOK SHEET <"+numRowsSrcWorkbook+">  WITH TARGET WORKBOOK SHEET <"+numRowsTgtWorkBook+"> ");
	       
	        /// Lets Display the contents in a format acceptable
	        
	     //   Sheet srcSheet ;
	       
	  ///      Sheet srcSheet = sourceWorkbook.getSheetAt(i);
	        System.out.println(srcSheet.getSheetName());
	        Iterator<Row> iterator = srcSheet.iterator();
	        int j=0;
	        while (iterator.hasNext()) {
	        	System.out.print("Row <"+ ++j +">");
	        	int k=0;
	            Row nextRow = iterator.next();
	            Iterator<Cell> cellIterator = nextRow.cellIterator();
	            ReusableBooks aBook = new ReusableBooks();
	    
	            while (cellIterator.hasNext()) {
	            k++;	
	            ///System.out.print("<");
	                Cell nextCell = cellIterator.next();
	                int columnIndex = nextCell.getColumnIndex();
	             ///   System.out.println(getCellValue(nextCell));;
	                switch (columnIndex) {
	                case 0:
	                	
	                    aBook.setColumn1((String) getCellValue(nextCell));
	                    System.out.print( getCellValue(nextCell) + ">");
	                    break;
	                case 1:
	                    aBook.setColumn2((String) getCellValue(nextCell));
	                    System.out.print( getCellValue(nextCell) + ">");
	                    break;
	                case 2:
	                    aBook.setColumn3((String) getCellValue(nextCell));
	                    System.out.print( getCellValue(nextCell) + ">");
	                    break;
	                case 3:
	                    aBook.setColumn4((String) getCellValue(nextCell));
	                    System.out.print( getCellValue(nextCell) + ">");
	                    break;
	                case 4:
	                    aBook.setColumn5((String) getCellValue(nextCell));
	                    System.out.print( getCellValue(nextCell) + ">");
	                    break;
	                }
	     
	     
	            }
	            System.out.println("@@@@");
	            listBooks.add(aBook);
	        }
	     
	        sourceWorkbook.close();
	        sourceInputStream.close();
	       /// }
	        return listBooks;
	        }
	    
	
}
