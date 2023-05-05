package com.nilesh.springExcelExport.excel;

import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.nilesh.springExcelExport.model.Student;

public class UserExcelImporter {
	public List<Student> excelImport(String url){
		List<Student> listStudent=new ArrayList<>();
		long sid=0;
		String sname="";
		String saddress="";
		String scity="";
		String spin="";
		
		//url= "C:\\Users\\Asus\\Desktop\\student.xlsx";

		String excelFilePath=url;
		
		long start = System.currentTimeMillis();
		
		FileInputStream inputStream;
		try {
			inputStream = new FileInputStream(excelFilePath);
			Workbook workbook=new XSSFWorkbook(inputStream);
			Sheet firstSheet=workbook.getSheetAt(0);
			Iterator<Row> rowIterator=firstSheet.iterator();
			rowIterator.next();
			
			while(rowIterator.hasNext()) {
				Row nextRow = rowIterator.next();
				Iterator<Cell> cellIterator=nextRow.cellIterator();
				while(cellIterator.hasNext()) {
					Cell nextCell=cellIterator.next();
					int columnIndex=nextCell.getColumnIndex();
					switch (columnIndex) {
					case 0:
						sid=(long)nextCell.getNumericCellValue();
						System.out.println(sid);
						break;
					case 1:
						sname=nextCell.getStringCellValue();
						System.out.println(sname);
						break;
					case 2:
						saddress=nextCell.getStringCellValue();
						System.out.println(saddress);
						break;
					case 3:
						scity=nextCell.getStringCellValue();
						System.out.println(scity);
						break;
					case 4:
						spin=nextCell.getStringCellValue();
						System.out.println(spin);
						break;
					
					}
					listStudent.add(new Student(sid, sname, saddress, scity, spin));
				}
			}
			
			workbook.close();
			long end = System.currentTimeMillis();
			System.out.printf("Import done in %d ms\n", (end - start));
			
		} catch (Exception e) {
			// TODO Auto-generated catch block
			
			e.printStackTrace();
		}
		
		return listStudent;
	}

}
