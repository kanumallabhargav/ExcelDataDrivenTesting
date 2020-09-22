package demos;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReaderDemo {

	public static void main(String[] args) throws IOException 
	{
		FileInputStream fis = new FileInputStream("/Users/bhargavkanmalla/Documents/selenium-java-3.141.59/FrameworkIdeas/KeyWordDrivenFramework/Demos/Book.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		
		int sheetCount = workbook.getNumberOfSheets();
		
		for (int i=0;i<sheetCount;i++)
		{
			if(workbook.getSheetName(i).equalsIgnoreCase("sheet1"))
			{
				XSSFSheet sheet = workbook.getSheetAt(i);
				Iterator<Row> rows= sheet.iterator();
				Row firstRow = rows.next();
				Iterator<Cell> cellItr = firstRow.cellIterator();
				int k=0;
				int column=0;
				while(cellItr.hasNext())
				{
					Cell value = cellItr.next();
					String valueInCell = value.getStringCellValue();
					if(valueInCell.equalsIgnoreCase("xls"))
					{	
						column=k;
					}
					k++;
				}
				while(rows.hasNext())
				{
					Row rw = rows.next();
					if(rw.getCell(column).getStringCellValue().equalsIgnoreCase("ds"))
					{
						Iterator<Cell> it = rw.cellIterator();
						while (it.hasNext())
						{
							String val = it.next().getStringCellValue();
							System.out.println(val);
						}
					}
				}
			}
		}
		workbook.close();
	}

}
