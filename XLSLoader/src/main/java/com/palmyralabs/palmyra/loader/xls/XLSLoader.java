package com.palmyralabs.palmyra.loader.xls;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Collections;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.stream.Collectors;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.zitlab.palmyra.client.PalmyraClient;
import com.zitlab.palmyra.client.pojo.Tuple;

public class XLSLoader {

	private int noRecords = 0;
	private int errRecords = 0;
	private int ignoredRecords = 0;
	private static final SimpleDateFormat sdf = new SimpleDateFormat("dd-MM-yyyy");
	private int rowCount = 0;
	private XSSFSheet errorSheet;

	public void loadXls(PalmyraClient client, String mappingFile, File xlsFile, String ciType, int sheetStartIndex,
			int sheetEndIndex, int startRow, String outFile) throws Exception {
		MappingLoader defectMapping = new MappingLoader(mappingFile);
		List<Mapping> mappings = defectMapping.getMappings();

		for (int i = sheetStartIndex; i <= sheetEndIndex; i++) {
			loadFile(client, ciType, mappings, xlsFile, i, startRow, outFile);
		}
	}

	public void loadFile(PalmyraClient client, String ciType, List<Mapping> mappings, File xlsFile, int sheetIndex,
			int startRow, String outFile) throws IOException, InvalidFormatException {
		loadFile(client, ciType, mappings, xlsFile, sheetIndex, startRow, outFile, Collections.emptyMap());
	}

	public void loadFile(PalmyraClient client, String ciType, List<Mapping> mappings, File xlsFile, int sheetIndex,
			int startRow, String outFile, Map<String, Object> defaultProps) throws IOException, InvalidFormatException {
		noRecords = 0;
		errRecords = 0;
		rowCount = 0;

		String fileName = xlsFile.getName();
		System.out.println("=========================================================");
		System.out.println("Started processing the file " + fileName + " Sheet " + sheetIndex);
		Workbook workbook, errorbook = null;
		errorbook = new XSSFWorkbook();
		errorSheet = (XSSFSheet) errorbook.createSheet("error");

		try {
			workbook = new XSSFWorkbook(xlsFile);

		} catch (Throwable t) {
			FileInputStream fis = new FileInputStream(xlsFile);
			workbook = new HSSFWorkbook(fis);
		}

		int endRow = -1;

		int skipCount = 0;

		int currentRow = 0;

		int numCellCheck = mappings.size() / 4;

		DataFormatter formatter = new DataFormatter();

		Sheet sheet = workbook.getSheetAt(sheetIndex);

		Iterator<Row> rowIterator = sheet.iterator();
		while (rowIterator.hasNext()) {
			currentRow++;
			if (skipCount > 15) {
				break;
			}

			Tuple tuple = new Tuple();
			tuple.setType(ciType);
			Row row = rowIterator.next();

			if (startRow > currentRow)
				continue;

			if (row.getPhysicalNumberOfCells() < numCellCheck) {
				ignoredRecords++;
				skipCount++;
				continue;
			}

			tuple = mapRow(mappings, currentRow, formatter, tuple, row);

			if (null != tuple) {
				assignDefaultProps(tuple, defaultProps);
				saveData(tuple, client, mappings, currentRow);
				skipCount = 0;
			} else {
				ignoredRecords++;
				skipCount++;
			}
			if (endRow > 0 && endRow == currentRow)
				break;

		}
		if (errRecords > 0) {
			try (FileOutputStream outputStream = new FileOutputStream(getFileName(outFile, sheetIndex))) {
				errorbook.write(outputStream);
				errorbook.close();
			}
		} else {
			if (null != errorbook)
				errorbook.close();
		}

		workbook.close();
		System.out.println(noRecords + " records has been synced with the system from " + fileName);
		if (errRecords > 0) {
			System.out.println(errRecords + " has been ignored due to errors");
		}
		if (ignoredRecords > 0) {
			System.out.println(ignoredRecords + " has been ignored due less Cells");
		}
		System.out.println("=========================================================");
	}

	private void assignDefaultProps(Tuple tuple, Map<String, Object> defaultProps) {
		if (null != defaultProps) {
			for (Entry<String, Object> entry : defaultProps.entrySet()) {
				tuple.setAttribute(entry.getKey(), entry.getValue());
			}
		}
	}

	private Tuple mapRow(List<Mapping> mappings, int currentRow, DataFormatter formatter, Tuple tuple, Row row) {
		for (Mapping mapping : mappings) {
			String key = mapping.getProperty();
			if(null !=mapping.getStaticValue()){
				tuple.setAttribute(key,  mapping.getStaticValue());
				continue;
			}
			
			Cell cell = row.getCell(mapping.getColumn());
			if (null != cell) {
				Object value = getValue(formatter, mapping, cell);
				if (mapping.isMandatory() && null == value) {
					tuple = null;
					System.out
							.println("row " + currentRow + ", column " + mapping.getColumn() + " should be mandatory");
					break;
				}
				if (mapping.getDataType() == Mapping.BIT) {
					value = (null == value) ? false : (value.toString().toLowerCase().startsWith("y") ? true : false);
				}
				
				tuple.setRefAttribute(key, value);
			}
		}
		return tuple;
	}

	private Object getValue(DataFormatter formatter, Mapping mapping, Cell cell) {
		Object value;
		CellType cellType = cell.getCellType();
		switch (cellType) {
		case NUMERIC:
			if (DateUtil.isCellDateFormatted(cell, null)) {
				value = cell.getDateCellValue();
				value = sdf.format(value);
				break;
			} else {
				value = getCellValue(cell, mapping, formatter);
				break;
			}
		case BLANK:
			value = null;
			break;
		default:
			value = getCellValue(cell, mapping, formatter);
			break;
		}
		return value;
	}

	private String getFileName(String outFile, int sheetIndex) {
		int idx = outFile.lastIndexOf('.');
		if (idx > 0) {
			return outFile.substring(0, idx) + "_" + sheetIndex + "." + outFile.substring(idx + 1, outFile.length());
		} else
			return outFile + "_" + sheetIndex;
	}

	public static Object getCellValue(Cell cell, Mapping mapping, DataFormatter formatter) {
		Object value;
		int dt = mapping.getDataType();
		Object val = formatter.formatCellValue(cell).trim();
		if (null == val)
			return null;
		try {
			switch (dt) {
			case Mapping.INTEGER:
				value = Integer.parseInt(val.toString());
				break;
			case Mapping.DOUBLE:
				value = Double.parseDouble(val.toString());
				break;
			case Mapping.DATE: {
				if (val instanceof Date) {
					value = sdf.format(val);
				} else
					value = val;
				break;
			}

			default:
				value = val;
				break;
			}
		} catch (NumberFormatException nfe) {
			value = null;
		}
		return value;
	}

	public void saveData(Tuple tuple, PalmyraClient client, List<Mapping> mapping, int curRow) {
		try {
			{
				client.save(tuple);
				noRecords++;
			}
		} catch (Throwable e) {
			// System.out.println(tuple.getAttribute("raisedOn"));
			Row row = errorSheet.createRow(++rowCount);
			List<Object> attributes = tuple.getAttributes().values().stream().collect(Collectors.toList());
			int columnCount = 0;

			Cell zell = row.createCell(++columnCount);
			zell.setCellValue(curRow);

			for (Object field : attributes) {
				Cell cell = row.createCell(++columnCount);
				if (field instanceof String) {
					cell.setCellValue((String) field);
				} else if (field instanceof Integer) {
					cell.setCellValue((Integer) field);
				}
			}

			Cell cell = row.createCell(++columnCount);
			cell.setCellValue(e.getMessage());

			for (Entry<String, Object> entry : tuple.getAttributes().entrySet()) {
				System.out.print(entry.getKey() + ":" + entry.getValue() + " ");
			}
			System.out.println();
			System.out.println("Error while processing record " + tuple.getAttributeAsString("defectId") + " error:"
					+ e.getMessage());
			errRecords++;
		}
	}

}
