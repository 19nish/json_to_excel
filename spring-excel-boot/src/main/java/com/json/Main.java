package com.json;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

public class Main {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		
		ObjectMapper om = new ObjectMapper();
		try {
			JsonNode node = om.readTree(new File("C:/Users/004DIG744/Documents/input.json"));
			JsonNode header = node.get("header");
			Iterator<JsonNode> it = header.iterator();
			XSSFWorkbook wb = new XSSFWorkbook();
			XSSFSheet sheet = wb.createSheet("Employee Details");
			Row row = sheet.createRow(0);
			int colNum = 0;
			while(it.hasNext()) {
				Cell cell = row.createCell(colNum);
				cell.setCellValue(it.next().asText());
			}
			JsonNode body =node.get("body");
			int rowNum = 1;
			colNum = 0;
			int i = 0;
			JsonNode rowNode;
			while(i < body.size()) {
				rowNode = body.get(i++);
				Row bodyRow = sheet.createRow(rowNum++);
				Cell nameCell = bodyRow.createCell(colNum++);
				Cell ageCell = bodyRow.createCell(colNum++);
				Cell depCell = bodyRow.createCell(colNum++);
				Cell salCell = bodyRow.createCell(colNum++);
				Cell manCell = bodyRow.createCell(colNum++);
				nameCell.setCellValue(rowNode.get("name").asText());
				ageCell.setCellValue(rowNode.get("age").asInt());
				depCell.setCellValue(rowNode.get("department").asText());
				salCell.setCellValue(rowNode.get("salary").asInt());
				manCell.setCellValue(rowNode.get("isManager").asBoolean());
				colNum = 0;
			}
			FileOutputStream outputStream = new FileOutputStream("C:/Users/004DIG744/Documents/test.xlsx");
			wb.write(outputStream);
			wb.close();
			System.out.println("Excel File Generated");
		}catch(JsonProcessingException e1){
			e1.printStackTrace();
		}catch (IOException e1) {
			e1.printStackTrace();	
		}finally {
			System.out.println("Code executed");
		}

	}

}
