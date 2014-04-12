package com.utility;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.TreeMap;

import org.apache.poi.hpsf.CustomProperties;
import org.apache.poi.hssf.dev.FormulaViewer;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.formula.eval.ErrorEval;
import org.apache.poi.ss.formula.eval.EvaluationException;
import org.apache.poi.ss.formula.eval.NumberEval;
import org.apache.poi.ss.formula.eval.OperandResolver;
import org.apache.poi.ss.formula.eval.ValueEval;
import org.apache.poi.ss.formula.functions.Function;
import org.apache.poi.ss.formula.functions.NumericFunction;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;


public class ExcelFormat  extends TreeMap<String, List<Double>>
{
	public static ExcelFormat map=new ExcelFormat();
	public static  ExcelFormat map1=new ExcelFormat();
	public  static ExcelFormat map2=new ExcelFormat();
	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;

	public void put(String key, Double number) {
        List<Double> current = get(key);
        if (current == null) {
            current = new ArrayList<Double>();
            super.put(key, current);
        }
        current.add(number);
    }

	public static void copyRawExcel(String source, String destination, boolean isRound1)throws IOException 
	{
		int columnNum=0;
		if(isRound1==true)
		{
			map=map1;
		}
		else
		{
			map=map2;
		}
		//ArrayList<Integer> list=new ArrayList<>();
		File f1;
		FileInputStream file1;
		
		try
		{
			f1=new File(source);
			file1=new FileInputStream(f1);
			Workbook workbook1=new HSSFWorkbook(file1);
		    Sheet sheet1=workbook1.getSheetAt(0);
		    
		    //This piece of code is reading value from the sorted excel and puting it in map
		    Iterator<Row> rowIterator1 = sheet1.iterator();
		    while(rowIterator1.hasNext())
		    {
		        Row row1 = rowIterator1.next();
		        Iterator<Cell> cellIterator1 = row1.cellIterator();
		        while(cellIterator1.hasNext())
		        {
		            Cell cell = cellIterator1.next();
		             switch(cell.getCellType()) 
		            {
	                	case Cell.CELL_TYPE_STRING:
	                   			Cell cell2=row1.getCell(1);
	                			double cellValue= cell2.getNumericCellValue();
	                			map.put(cell.getStringCellValue(), cellValue);
	                			break;
		            }
		        }  
		    }
		    file1.close();
		    
		    //This piece of code will take the keyset of map and write it into excel 
		    HSSFWorkbook workbook = new HSSFWorkbook();
		    HSSFSheet sheet = workbook.createSheet("Sheet 1");
		
		    //Create a new row in current sheet
		    Row row = sheet.createRow(0);
		   // System.out.println("MAp:"+map);
		    //Create a new cell in current row
		    for(String key:map.keySet())
		    {
		    	
		    	if(key.contains("Login")||key.contains("Logout")||key.contains("New_Debt_Instr")||key.contains("New_Program_Header")||key.contains("Create_New_Program")||key.contains("Create_New_Equity")||key.contains("Create_Watch_List_Entry")||key.contains("Create_Debt_Copy")||key.contains("Create_Organization_Merge")||key.contains("Create_CFG_Analyst_Profile"))
		    	{
		    	}
		    	else
		    	{
		    		Cell cell1=row.createCell(columnNum);
			    	cell1.setCellValue(key);
			    	List<Double> columnValues = map.get(key);
			    	int temp=1;
			    	for(double i:columnValues)
			    	{
			    		Row row2;
			    		if(sheet.getPhysicalNumberOfRows()-1>temp-1)
			    		{
			    			row2=sheet.getRow(temp);
			       		}
			    		else
			    		{
			    			row2=sheet.createRow(temp);
			    		}
			    		Cell cell2=row2.createCell(columnNum);
		    			cell2.setCellValue(i);
			    		temp=temp+1;
			    	}
			    	columnNum=columnNum+1;
			    }
		    }
		    try 
		    {
		    	FileOutputStream out;
		      	if(isRound1==true)
		    	{
		      		out=new FileOutputStream(new File(ExcelUtilitySwing.destinationDir+"\\round1_temp.xls"));
		    	}
		      	else
		      	{
		      		out=new FileOutputStream(new File(ExcelUtilitySwing.destinationDir+"\\round2_temp.xls"));
		      	}
				workbook.write(out);
			    out.close();
			    calculateFormulae(isRound1);
		    }
		    catch (FileNotFoundException e)
		    {
		    	System.out.println(e.getMessage());
		        e.printStackTrace();
		    }
			
		}
		catch(Exception e)
		{
			e.printStackTrace();
			AlertBox.alert("File not found!!!");
		}
	}
	public static void calculateFormulae(boolean isRound1)
	{
		try
		{
			 //This piece of code is calculating average and 90%, it is taking help of map to find out the row in which it has to write
			FileInputStream file1;
			if(isRound1==true)
			{
				map=map1;
				file1=new FileInputStream(ExcelUtilitySwing.destinationDir+"\\round1_temp.xls");
			}
			else
			{
				map=map2;
				file1=new FileInputStream(ExcelUtilitySwing.destinationDir+"\\round2_temp.xls");
			}
			CellStyle style;
			int columnIndex=0;
			HSSFWorkbook workbook1=new HSSFWorkbook(file1);
			
			 
			Sheet sheet1=workbook1.getSheetAt(0);
		 	Row row1 = sheet1.createRow(sheet1.getPhysicalNumberOfRows());
		    Row row2 = sheet1.createRow(sheet1.getPhysicalNumberOfRows());
		    style = workbook1.createCellStyle();
		    style.setFillForegroundColor(HSSFColor.YELLOW.index);
		    style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		    for(String key:map.keySet())
		    {
		    	//System.out.println(map.get(key));
		    	if(key.contains("Login")||key.contains("Logout")||key.contains("New_Debt_Instr")||key.contains("New_Program_Header")||key.contains("Create_New_Program")||key.contains("Create_New_Equity")||key.contains("Create_Watch_List_Entry")||key.contains("Create_Debt_Copy")||key.contains("Create_Organization_Merge")||key.contains("Create_CFG_Analyst_Profile"))
		    	{
		    	}
		    	else
		    	{
			    	List<Double> columnValues = map.get(key);
			    	//System.out.println(columnValues);
			    	//System.out.println("Range:"+columnValues.size());
			    	//System.out.println(sheet1.getPhysicalNumberOfRows());
			    	
			    	//this piece of code is getting a blank row, creating a cell, setting cell average formulae and setting style of yellow color
			       	Cell cell1=row1.createCell(columnIndex);
			    	 cell1.setCellType(HSSFCell.CELL_TYPE_FORMULA);
			    	
			    	 cell1.setCellStyle(style);
			    	 CellReference ref = new CellReference(0,columnIndex);
			    	 String columnIndexName=ref.formatAsString();
			    	 if(columnIndexName.length()>4)
			    	 {
			    		 columnIndexName=columnIndexName.substring(1, columnIndexName.length()-2);		    	 
				   	 }
			    	 else
			    	 {
			    		 columnIndexName=columnIndexName.substring(1, 2);		    	 
				    }
			    	 cell1.setCellFormula("AVERAGE("+columnIndexName+"2:"+columnIndexName+""+(columnValues.size()+1)+")");
	
			    	//this piece of code is getting a blank row, creating a cell, setting cell 90% formulae and setting style of yellow color
			       	 Cell cell2=row2.createCell(columnIndex);
			    	 cell2.setCellType(HSSFCell.CELL_TYPE_FORMULA);
			    	 cell2.setCellStyle(style);
			    	 cell2.setCellFormula("PERCENTILE("+columnIndexName+"2:"+columnIndexName+""+(columnValues.size()+1)+",0.9)");
			    	  
			    	 columnIndex++;
		    	}
		    	
		    	file1.close();
		    }
		    try 
		    {
		    	FileOutputStream out;
		    	if(isRound1==true)
				{
					out=new FileOutputStream(new File(ExcelUtilitySwing.destinationDir+"\\round1_temp.xls"));
				}
				else
				{
					out=new FileOutputStream(new File(ExcelUtilitySwing.destinationDir+"\\round2_temp.xls"));
				}
		        workbook1.write(out);
		        out.close();
		    }
		    catch (FileNotFoundException e)
		    {
		    	System.out.println(e.getMessage());
		        e.printStackTrace();
		    }
		 }
		catch(Exception e)
		{
			System.out.println(e.getMessage());
			e.printStackTrace();
		}
	}
	void excelCopy(String lookUp,String template,String result)
	{
		FileInputStream file1;
		int lastRowUpdated=0,currentRow=0;
	
		try 
		{
			file1 = new FileInputStream(lookUp);
			int columnIndex=0;
			Workbook lookupWorkbook=new HSSFWorkbook(file1);
			Sheet lookupSheet=lookupWorkbook.getSheetAt(0);
			
			FileInputStream file2=new FileInputStream(ExcelUtilitySwing.destinationDir+"\\round1_temp.xls");
			HSSFWorkbook round1Workbook=new HSSFWorkbook(file2);
			FormulaEvaluator evaluator = round1Workbook.getCreationHelper().createFormulaEvaluator();
			
			Sheet round1Sheet=round1Workbook.getSheetAt(0);
			
			FileInputStream file3=new FileInputStream(template);
			Workbook templateWorkbook=new HSSFWorkbook(file3);
			Sheet templateSheet=templateWorkbook.getSheetAt(0);
			
			FileInputStream file4=new FileInputStream(ExcelUtilitySwing.destinationDir+"\\round2_temp.xls");
			HSSFWorkbook round2Workbook=new HSSFWorkbook(file4);
			FormulaEvaluator evaluator2 = round2Workbook.getCreationHelper().createFormulaEvaluator();
			Sheet round2Sheet=round2Workbook.getSheetAt(0);
						
			int averageRow=round1Sheet.getPhysicalNumberOfRows()-2;
			int percentileRow=round1Sheet.getPhysicalNumberOfRows()-1;
		
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
			        			evaluator.evaluateFormulaCell(round1CellAverage);
			        			evaluator2.evaluateFormulaCell(round2CellAverage);
			        			templateCellFormulae1.setCellFormula("IF(B"+(currentRow+1)+">D"+(currentRow+1)+",((B"+(currentRow+1)+"-D"+(currentRow+1)+")/B"+(currentRow+1)+")*100,((B"+(currentRow+1)+"-D"+(currentRow+1)+")/D"+(currentRow+1)+")*100)");
			        			//System.out.println(round1CellAverage.getNumericCellValue());
			        			templateCellRound1Average.setCellValue(round1CellAverage.getNumericCellValue());
			        			templateCellRound2Average.setCellValue(round2CellAverage.getNumericCellValue());
			        		
			        			
			        			Row round1RowPercentile=round1Sheet.getRow(percentileRow);
			        			Cell round1CellPercentile=round1RowPercentile.getCell(columnIndex);
			        			Row round2RowPercentile=round2Sheet.getRow(percentileRow2);
			        			Cell round2CellPercentile=round2RowPercentile.getCell(columnIndex);
			        			Cell templateCellRound1Percentile=templateRow.getCell(2);
			        			Cell templateCellRound2Percentile=templateRow.getCell(4);
			        			evaluator.evaluateFormulaCell(round1CellPercentile);
			        			evaluator2.evaluateFormulaCell(round2CellPercentile);
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
			
			FileOutputStream out =new FileOutputStream(ExcelUtilitySwing.resultFileName);
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


