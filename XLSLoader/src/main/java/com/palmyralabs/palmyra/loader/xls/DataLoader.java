package com.palmyralabs.palmyra.loader.xls;

import java.io.File;
import java.io.IOException;
import java.util.List;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import com.zitlab.palmyra.client.PalmyraClient;

public abstract class DataLoader {

	private XLSLoader loader;
	private PalmyraClient client;

	public DataLoader(PalmyraClient client, XLSLoader loader) {
		this.loader = loader;
		this.client = client;
	}

	public abstract void loadData();
	
	public void loadXls(String mappingFile, String xlsFileName, String ciType, int sheetStartIndex, int sheetEndIndex,
			int startRow, String outFile, Map<String, Object> defaultProps) {
		MappingLoader defectMapping = new MappingLoader(mappingFile);
		List<Mapping> mappings = defectMapping.getMappings();

		File xlsFolder = new File(xlsFileName);
		if (!xlsFolder.exists()) {
			throw new RuntimeException("file " + xlsFileName + " not found in the system");
		}
		for (int i = sheetStartIndex; i <= sheetEndIndex; i++) {
			try {
				loader.loadFile(client, ciType, mappings, xlsFolder, i, startRow, outFile, 
						defaultProps);
			} catch (InvalidFormatException | IOException e) {
				e.printStackTrace();
				throw new RuntimeException("Sheet " + i + " cannot be processed", e);				
			}
		}
	}
}
