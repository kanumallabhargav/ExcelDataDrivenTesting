import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class AccessTbl 
{	
	public ArrayList<String> getData(String testCaseName) throws IOException
	{
		//get access to the excel file
				FileInputStream fis = new FileInputStream("/Users/bhargavkanmalla/Documents/selenium-java-3.141.59/TC_Sheet.xlsx");
				XSSFWorkbook workbook = new XSSFWorkbook(fis);
				
				ArrayList<String> al = new ArrayList<String>();
				int sheets = workbook.getNumberOfSheets();
				for (int i=0;i<sheets;i++)
				{
					if(workbook.getSheetName(i).equalsIgnoreCase("Sheet 1"))
					{
						//Access the sheet
						XSSFSheet sheet = workbook.getSheetAt(i);
						
						//scanning the first row to find the ExpResult row
						Iterator<Row> rows=sheet.iterator();
						Row firstrow = rows.next();
						Iterator<Cell> ce=firstrow.iterator();
						int col=0, grab=0;
						while (ce.hasNext())
						{
							Cell value=ce.next();
							if(value.getStringCellValue().equalsIgnoreCase("testcasename"))
							{
								col=grab;
							}
							grab++;
						}
						System.out.println(col);
						
						while(rows.hasNext())
						{
							Row rw = rows.next();
							if(rw.getCell(col).getStringCellValue().equalsIgnoreCase(testCaseName))
							{
								Iterator <Cell> cv = rw.cellIterator();
								while(cv.hasNext())
								{
									Cell c=cv.next();
									if(c.getCellTypeEnum()==CellType.STRING)
                                    {
    									al.add(c.getStringCellValue());
                                    }
									else
									{
										al.add(NumberToTextConverter.toText(c.getNumericCellValue()));
										
									}
								}
							}
						}
					}
				}
				return al;
	}
}
