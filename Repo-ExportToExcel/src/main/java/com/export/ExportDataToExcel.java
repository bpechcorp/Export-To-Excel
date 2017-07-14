package com.export;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.alfresco.model.ContentModel;
import org.alfresco.repo.content.MimetypeMap;
import org.alfresco.service.cmr.repository.ChildAssociationRef;
import org.alfresco.service.cmr.repository.NodeRef;
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
		String nodeRefParam=req.getParameter("nodeRef");
		if(nodeRefParam!=null){
			
			writeDataInExcel(res,new NodeRef(nodeRefParam));
		}
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
		listString.add("Name");
		listString.add("Title");
		listString.add("Description");
		listString.add("Created Date");
		return listString;
	}

	/**
	 * This method is used for creating multiple row.So the inner list object
	 * will contain the cell details and the outer one will contain row details
	 * 
	 * @return List of values which we need to write in excel
	 */
	public List<List<Object>> getData(NodeRef folderNode) {
		List<List<Object>> listData = new ArrayList<>();
		
		List<ChildAssociationRef> childAssociationRefList=nodeService.getChildAssocs(folderNode);
		for(ChildAssociationRef child:childAssociationRefList){
			List<Object> row = new ArrayList<>();
			row.add(nodeService.getProperty(child.getChildRef(),ContentModel.PROP_NAME));
			row.add(nodeService.getProperty(child.getChildRef(),ContentModel.PROP_TITLE));
			row.add(nodeService.getProperty(child.getChildRef(),ContentModel.PROP_DESCRIPTION));
			row.add(nodeService.getProperty(child.getChildRef(),ContentModel.PROP_CREATED));
			listData.add(row);
		}
		return listData;
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
	public void writeDataInExcel(WebScriptResponse res,NodeRef folderNode) throws IOException {
		// Create Work Book and WorkSheet
		Workbook wb = new XSSFWorkbook();
		Sheet sheet = wb.createSheet("ExcelFile");
		sheet.createFreezePane(0, 1);

		generateHeaderInExcel(sheet);
		generateDataInExcel(sheet,folderNode);
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
	public void generateDataInExcel(Sheet sheet,NodeRef folderNode) {
		List<List<Object>> listOfData = getData(folderNode);

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
				} else if (obj instanceof Date) {
					c.setCellValue(obj.toString());
				}
				colNum++;
			}
			rowNum++;
		}
	}
}
