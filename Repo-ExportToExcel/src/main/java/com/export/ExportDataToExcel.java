package com.export;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.alfresco.repo.content.MimetypeMap;
import org.alfresco.service.cmr.repository.NodeService;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.extensions.webscripts.AbstractWebScript;
import org.springframework.extensions.webscripts.WebScriptRequest;
import org.springframework.extensions.webscripts.WebScriptResponse;

public class ExportDataToExcel extends AbstractWebScript {
	private String fileName = "ExcelSheet";
	private String format = "xlsx";
	
	private NodeService nodeService;
	
	public NodeService getNodeService() {
		return nodeService;
	}

	public void setNodeService(NodeService nodeService) {
		this.nodeService = nodeService;
	}

	/**
	 * This method is inherited from AbstractWebscript class and called
	 * internally.
	 */
	@Override
	public void execute(WebScriptRequest req, WebScriptResponse res) throws IOException {
		writeDataInExcel(res);
	}

	/**
	 * 
	 * @param res
	 *            Used for controlling the response.In our case we want to give
	 *            file as response.
	 * @param data
	 *            Content which we need to write in excel file
	 * @throws IOException
	 */
	public void ResponseDataInExcel(WebScriptResponse res, byte[] data) throws IOException {
		String filename = this.fileName + "." + this.format;
		// Set Header for downloading file.
		res.addHeader("Content-Disposition", "attachment; filename=" + filename);
		// Set spredsheet data
		byte[] spreadsheet = null;
		res.setContentType(MimetypeMap.MIMETYPE_OPENXML_SPREADSHEET);
		res.getOutputStream().write(data);
	}

	/**
	 * 
	 * @return List of header which we need to write in excel
	 */
	public List<String> getHeaderList() {
		List<String> listString = new ArrayList<>();
		listString.add("Header1");
		listString.add("Header2");
		listString.add("Header3");
		listString.add("Header4");
		return listString;
	}

	/**
	 * This method is used for creating multiple row.So the inner list object
	 * will contain the cell details and the outer one will contain row details
	 * 
	 * @return List of values which we need to write in excel
	 */
	public List<List<Object>> getData() {
		List<List<Object>> listString = new ArrayList<>();
		// Adding sample data
		// ******************
		// Creating sample Row 1
		List<Object> sampleRow1 = new ArrayList<>();
		sampleRow1.add("sampleCell11");
		sampleRow1.add("sampleCell12");
		sampleRow1.add("sampleCell13");
		sampleRow1.add("sampleCell14");

		// Creating sample Row 2
		List<Object> sampleRow2 = new ArrayList<>();
		sampleRow2.add(new Employee("Jhon"));
		sampleRow2.add(new Employee("Ethan"));
		sampleRow2.add(new Employee("Kevin"));
		sampleRow2.add(new Employee("Mike"));

		// Adding sample row in row list
		listString.add(sampleRow1);
		listString.add(sampleRow2);

		return listString;
	}

	/**
	 * 
	 * This method is used for Writing data in byte array.It is calling other
	 * methods as well which will help to generate data.
	 * 
	 * @param res
	 *            Writing content in output stream
	 * @throws IOException
	 */
	public void writeDataInExcel(WebScriptResponse res) throws IOException {
		// Create Work Book and WorkSheet
		Workbook wb = new XSSFWorkbook();
		Sheet sheet = wb.createSheet("ExcelFile");
		sheet.createFreezePane(0, 1);

		generateHeaderInExcel(sheet);
		generateDataInExcel(sheet);
		ByteArrayOutputStream baos = new ByteArrayOutputStream();
		wb.write(baos);
		ResponseDataInExcel(res, baos.toByteArray());
	}

	/**
	 * Used for generating header in excel.
	 * 
	 * @param sheet
	 *            It is an excel sheet object, where the header will be created
	 */
	public void generateHeaderInExcel(Sheet sheet) {
		List<String> headerValues = getHeaderList();
		Row hr = sheet.createRow(0);
		for (int i = 0; i < headerValues.size(); i++) {
			Cell c = hr.createCell(i);
			c.setCellValue(headerValues.get(i));
		}
	}

	/**
	 * Used for generating data in excel sheet
	 * 
	 * @param sheet
	 *            sheet It is an excel sheet object, where the data will be
	 *            written
	 */
	public void generateDataInExcel(Sheet sheet) {
		List<List<Object>> listOfData = getData();

		// Give first row as 1 and column as 0
		int rowNum = 1, colNum = 0;
		for (List<Object> rowValues : listOfData) {
			Row r = sheet.createRow(rowNum);
			colNum = 0;
			for (Object obj : rowValues) {
				Cell c = r.createCell(colNum);
				// Here you can add n number of condition for identifying the
				// type of object and based on that fetch value of it
				if (obj instanceof String) {
					c.setCellValue(obj.toString());
				} else if (obj instanceof Employee) {
					c.setCellValue(((Employee) obj).getName());
				}
				colNum++;
			}
			rowNum++;
		}
	}
}
