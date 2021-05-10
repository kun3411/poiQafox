package poi;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Demotest {
	public static void main(String[] args) throws IOException {
		ArrayList<String> alist = getdatafromexcelfile("register");
		for(String a:alist) {
			System.out.println(a);
		}
	}

	public static ArrayList<String> getdatafromexcelfile(String Testname) throws IOException {
		ArrayList<String> alist=new ArrayList<String>();
		FileInputStream fis=new FileInputStream("C:\\Users\\Yogi\\Desktop\\ExcelTestData.xlsx");
		XSSFWorkbook workbook=new XSSFWorkbook(fis);
		int sheetcount=workbook.getNumberOfSheets();
		for(int i=0;i<sheetcount;i++) {
			if(workbook.getSheetName(i).equalsIgnoreCase("sheetA")) {
				XSSFSheet sheet=workbook.getSheetAt(i);
				Iterator<Row> rows=sheet.iterator();
				
				Row firstrow=rows.next();
				Iterator<Cell> firstrowcells = firstrow.iterator();
				int c=0;
				int testcolumnposition=0;
				
				while(firstrowcells.hasNext()) {
					//System.out.println(firstrowcells.next().getStringCellValue());
	             			Cell firstrowcell = firstrowcells.next();
	             			if(firstrowcell.getStringCellValue().equalsIgnoreCase("Tests")) {
	             				testcolumnposition=c;
	             				
	             			}
	             			c++;
	             			while(rows.hasNext()) {
	             				Row row = rows.next();
	             				Cell cell = row.getCell(testcolumnposition);
	             				if(cell.getStringCellValue().equals("Register")) {
	             			      Iterator<Cell> cells = row.iterator();
	             			      cells.next();
	             			      while(cells.hasNext()) {
	             			    	  //System.out.println(cells.next().getStringCellValue());
	             			    	  //when cell has numeric value
	             			    	  Cell currentcell = cells.next();
	             			    	  if(currentcell.getCellType()==CellType.STRING) {
	             			    		  //System.out.println(currentcell.getStringCellValue());
	             			    		  alist.add(currentcell.getStringCellValue());
	             			    	  }
	             			    	  else if(currentcell.getCellType()==CellType.NUMERIC) {
	             			    		 //System.out.println(NumberToTextConverter.toText(currentcell.getNumericCellValue()));
	             			    		  //System.out.println(currentcell.getNumericCellValue());
	             			    		  alist.add(NumberToTextConverter.toText(currentcell.getNumericCellValue()));
	             			    	  }
	             			      }
	             				}
	             				
	             			}
	             			
				}
				
			}
		}
		return alist;
		

	}

}
