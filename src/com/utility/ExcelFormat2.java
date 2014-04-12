package com.utility;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class ExcelFormat2 
{
	public static void main(String arg[])
	{
		ExcelFormat2 a=new ExcelFormat2();
		a.excelCopy();		
	}
	void excelCopy()
	{
		FileInputStream file1;
		int lastRowUpdated=0,currentRow=0;
		int i=0;
		try 
		{
			file1 = new FileInputStream("C:\\Users\\SharvaP\\Desktop\\Excel Format\\LookUp for Transaction Name.xls");
			int columnIndex=0;
			Workbook lookupWorkbook=new HSSFWorkbook(file1);
			Sheet lookupSheet=lookupWorkbook.getSheetAt(0);
			
			FileInputStream file2=new FileInputStream("C:\\Users\\SharvaP\\Desktop\\Excel Format\\round1_temp.xls");
			Workbook round1Workbook=new HSSFWorkbook(file2);
			Sheet round1Sheet=round1Workbook.getSheetAt(0);
			
			FileInputStream file3=new FileInputStream("C:\\Users\\SharvaP\\Desktop\\Excel Format\\template.xls");
			Workbook templateWorkbook=new HSSFWorkbook(file3);
			Sheet templateSheet=templateWorkbook.getSheetAt(0);
			
			FileInputStream file4=new FileInputStream("C:\\Users\\SharvaP\\Desktop\\Excel Format\\round2_temp.xls");
			Workbook round2Workbook=new HSSFWorkbook(file4);
			Sheet round2Sheet=round2Workbook.getSheetAt(0);
						
			int averageRow=round1Sheet.getPhysicalNumberOfRows()-2;
			int percentileRow=round1Sheet.getPhysicalNumberOfRows()-1;
		System.out.println("Average row1:"+averageRow);
			int averageRow2=round2Sheet.getPhysicalNumberOfRows()-2;
			int percentileRow2=round2Sheet.getPhysicalNumberOfRows()-1;
		
			/*Row round1Row = round1Sheet.getRow(averageRow);
			Row round2Row = round2Sheet.getRow(averageRow2);
			
			Cell tempRound1Cell=round1Row.getCell(columnIndex);
			Cell tempRound2Cell=round2Row.getCell(columnIndex);
			CellStyle cellStyle = templateWorkbook.createCellStyle();

			CreationHelper createHelper=templateWorkbook.getCreationHelper();
			cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("#"));
			*/
			
			Iterator<Row> rowIterator1 = lookupSheet.iterator();
		    while(rowIterator1.hasNext())
		    {
		    	
		        Row lookupRow = rowIterator1.next();
		        if(lookupRow.getRowNum()>0)
		        {
		        	Cell lookupCell=lookupRow.getCell(1);
		        	String transationNameTemplate=lookupCell.getStringCellValue();
		        	//System.out.println("NAme:"+transationNameTemplate);
		        	//System.out.println(templateSheet.getPhysicalNumberOfRows());
		        	for(int count=lastRowUpdated;count<templateSheet.getPhysicalNumberOfRows()-1;count++)
		        	{
		        		Row templateRow=templateSheet.getRow(count);
		        	
		        		Cell templateCell=templateRow.getCell(0);
		        		Cell templateCellFormulae1=templateRow.getCell(5);
		        		Cell templateCellFormulae2=templateRow.getCell(6);
		        		
		        		String templateCellName=templateCell.getStringCellValue();
		        		//System.out.println("template cell name:"+templateCellName);
		        		currentRow=templateRow.getRowNum();
		        		if(templateCellName.equals(transationNameTemplate))
		        		{
		        			
		        				//System.out.println("Transac Name:"+templateCellName+", row num:"+currentRow+", last updated row:"+lastRowUpdated);
			        			Row round1RowAverage=round1Sheet.getRow(averageRow);
			        			Cell round1CellAverage=round1RowAverage.getCell(columnIndex);
			        			Row round2RowAverage=round2Sheet.getRow(averageRow2);
			        			Cell round2CellAverage=round2RowAverage.getCell(columnIndex);
			        			Cell templateCellRound1Average=templateRow.getCell(1);
			        			Cell templateCellRound2Average=templateRow.getCell(3);
			        			templateCellFormulae1.setCellFormula("IF(B"+(currentRow+1)+">D"+(currentRow+1)+",((B"+(currentRow+1)+"-D"+(currentRow+1)+")/B"+(currentRow+1)+")*100,((B"+(currentRow+1)+"-D"+(currentRow+1)+")/D"+(currentRow+1)+")*100)");
			        			System.out.println(round1CellAverage.getNumericCellValue());
			        			templateCellRound1Average.setCellValue(round1CellAverage.getNumericCellValue());
			        			templateCellRound2Average.setCellValue(round2CellAverage.getNumericCellValue());
			        		
			        			
			        			Row round1RowPercentile=round1Sheet.getRow(percentileRow);
			        			Cell round1CellPercentile=round1RowPercentile.getCell(columnIndex);
			        			Row round2RowPercentile=round2Sheet.getRow(percentileRow2);
			        			Cell round2CellPercentile=round2RowPercentile.getCell(columnIndex);
			        			Cell templateCellRound1Percentile=templateRow.getCell(2);
			        			Cell templateCellRound2Percentile=templateRow.getCell(4);
			        			templateCellFormulae2.setCellFormula("IF(C"+(currentRow+1)+">E"+(currentRow+1)+",((C"+(currentRow+1)+"-E"+(currentRow+1)+")/C"+(currentRow+1)+")*100,((C"+(currentRow+1)+"-E"+(currentRow+1)+")/E"+(currentRow+1)+")*100)");
			        			templateCellRound1Percentile.setCellValue(round1CellPercentile.getNumericCellValue());
			        			templateCellRound2Percentile.setCellValue(round2CellPercentile.getNumericCellValue());
			        			
			        			//System.out.println("round1 avg value:"+round1CellAverage.getNumericCellValue());
			        			//System.out.println("round2 avg value:"+round2CellAverage.getNumericCellValue());
			        			
			        			//System.out.println("Match Found"+templateRow.getCell(1).getNumericCellValue());
			        			columnIndex++;
			        			lastRowUpdated=currentRow;
			        			break;
		        			}
		        	}
		        }
		      
       		}
			
			FileOutputStream out =new FileOutputStream("C:\\Users\\SharvaP\\Desktop\\Excel Format\\Template.xls");
	        templateWorkbook.write(out);
	        out.close();
	        System.out.println("Finished....");
	        file1.close();
	        file2.close();
	        file3.close();
	        file4.close();
		} 
		catch (FileNotFoundException e) 
		{
			System.out.println(e.getMessage());
			e.printStackTrace();
		} 
		catch (IOException e)
		{
			System.out.println(e.getMessage());
			e.printStackTrace();
		}
	
		
	}
	
}
