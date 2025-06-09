/**
 * 
 */
package mapObjects;

import java.io.FileInputStream;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import ui.DialogBoxUI;


/**
 * A data map file as an easily manipulatable Java object with helpful attributes
 * @author Hannah Layton
 */
public class IEDMap {
	/** Variable used to manipulate an excel file, in this case an IED Map, with Java */
	private XSSFWorkbook workbook;
	/** Device name listed on IED Map */
	private String deviceName;
	/** Relay Element column in IED Map */
	private int wordbitColumn = -1;
	/** RTAC Tag Alias column in IED Map */
	private int rtacAliasColumn = -1;
	/** Relay elements included in IED Map */
	private TreeMap<String, String> analogPoints;
	
	/**
	 * Converts a data map file into an easily manipulatable Java object with helpful attributes
	 * @param iedName - data map file stream to read from to add helpful attributes to the Java objects
	 */
	public IEDMap(FileInputStream iedName) {
		openIEDMap(iedName);
		setDeviceName();
		setWordbitColumn();
		setRtacAlias();
		setAnalogPoints();
	}

	/**
	 * Returns a workbook object representing the data map 
	 * @return a workbook object representing the data map
	 */
	public XSSFWorkbook getIEDMap() {
		return workbook;
	}

	/**
	 * Using the data map file stream, a workbook object is created to read the data map as an Excel file in Java
	 * @param iedName - data map file stream to read from to add helpful attributes to the Java objects
	 */
	public void openIEDMap(FileInputStream iedName) {
		try {
			this.workbook = new XSSFWorkbook(iedName);
		} catch (Exception e) {
			DialogBoxUI.infoBox("Could not open IED Map.", "");
			throw new IllegalArgumentException("Could not open IED Map.");
		}
	}

	/**
	 * Returns the name of the device in the data map
	 * @return the name of the device in the data map
	 */
	public String getDeviceName() {
		return deviceName;
	}

	/**
	 * Finds the name of the device in the data map & saves it to the Java object
	 */
	public void setDeviceName() {
		// Consider adding iedName parameter to indicate which data map does not have an Analog Points sheet
		if (workbook.getSheet("Analog Points") == null) {
			DialogBoxUI.infoBox("IED Map does not have Analog Points sheet.", "");
			throw new IllegalArgumentException("IED Map does not have Analog Points sheet.");
		}
		String deviceName = workbook.getSheet("Analog Points").getRow(2).getCell(1).getStringCellValue();
		if (deviceName != null && !deviceName.equals("")) {
			if (deviceName.startsWith("K") || deviceName.startsWith("L")) {
				this.deviceName = deviceName.substring(1);
			}
			else {
				this.deviceName = deviceName;
			}
		}
		else {
			DialogBoxUI.infoBox("Device Name in IED Map is not in B3.", "");
			throw new IllegalArgumentException("Device Name in IED Map is not in B3.");
		}
	}

	/**
	 * Returns the column that the wordbits in the data map are contained in
	 * @return the column that the wordbits in the data map are contained in
	 */
	public int getWordbitColumn() {
		return wordbitColumn;
	}

	/**
	 * Finds the column that the wordbits in the data map are contained in & saves it to the Java object
	 */
	public void setWordbitColumn() {
		for (Row row : workbook.getSheet("Analog Points")) {
			for (Cell cell : row) {
				if (cell.getCellType() == CellType.NUMERIC) {
					continue;
				}
				else if (cell.getStringCellValue().toLowerCase().contains("relay") && cell.getStringCellValue().toLowerCase().contains("element")) {
					this.wordbitColumn = cell.getColumnIndex();
					break;
				}
			}
			if (wordbitColumn >= 0) {
				break;
			}
		}
		if (wordbitColumn < 0) {
			DialogBoxUI.infoBox("Relay Element column could not be found in IED Map.", "");
			throw new IllegalArgumentException("Relay Element column could not be found in IED Map.");
		}
	}

	/**
	 * Returns RTAC tag alias for analog point
	 * @return RTAC tag alias for analog point
	 */
	public int getRtacAlias() {
		return rtacAliasColumn;
	}

	/**
	 * Finds the column that the RTAC tag aliases in the data map are contained in & saves it to the Java object
	 */
	public void setRtacAlias() {
		for (Row row : workbook.getSheet("Analog Points")) {
			for (Cell cell : row) {
				if (cell.getCellType() == CellType.NUMERIC) {
					continue;
				}
				else if (cell.getStringCellValue().toLowerCase().contains("rtac") && cell.getStringCellValue().toLowerCase().contains("name")) {
					this.rtacAliasColumn = cell.getColumnIndex();
					break;
				}
			}
			if (rtacAliasColumn >= 0) {
				break;
			}
		}
		if (rtacAliasColumn < 0) {
			DialogBoxUI.infoBox("RTAC Point Name column could not be found in IED Map.", "");
			throw new IllegalArgumentException("RTAC Point Name column could not be found in IED Map.");
		}
	}

	/**
	 * Returns a tree of entries in the data map with attributes for the wordbit & the RTAC alias
	 * @return a tree of entries in the data map with attributes for the wordbit & the RTAC alias
	 */
	public TreeMap<String, String> getAnalogPoints() {
		return analogPoints;
	}

	/**
	 * Adds entries in the data map to a tree map with attributes for the wordbit & the RTAC alias
	 */
	public void setAnalogPoints() {
		XSSFSheet currentSheet = workbook.getSheet("Analog Points");
		analogPoints = new TreeMap<String, String>();
		int row = 5;
		// Checks if data map is empty of wordbits
		if (currentSheet != null && currentSheet.getRow(row) != null && currentSheet.getRow(row).getCell(wordbitColumn) != null) {
			// Saves wordbit in row
			String currentRelayElement = currentSheet.getRow(row).getCell(wordbitColumn).getStringCellValue();
			// Loops through the rest of the entries in the data map
			while (currentSheet != null && currentSheet.getRow(row) != null && currentSheet.getRow(row).getCell(0) != null && !currentRelayElement.equals("")) {
				analogPoints.put(currentSheet.getRow(row).getCell(wordbitColumn).getStringCellValue(), currentSheet.getRow(row).getCell(rtacAliasColumn).getStringCellValue());
				row++;
				// If this was the last entry in the data map, the loop can end
				if (row > currentSheet.getLastRowNum() || currentSheet.getRow(row) == null || currentSheet.getRow(row).getCell(wordbitColumn) == null || currentSheet.getRow(row).getCell(wordbitColumn).getStringCellValue().equals("")) {
					break;
				}
				currentRelayElement = currentSheet.getRow(row).getCell(wordbitColumn).getStringCellValue();
			}
		}
		else {
			DialogBoxUI.infoBox("IED Map could not be read.", "");
			throw new IllegalArgumentException("IED Map could not be read.");
		}
		try {
			workbook.close();
		} catch (Exception e) {
			DialogBoxUI.infoBox("IED Map failed to close.", "");
			throw new IllegalArgumentException("IED Map failed to close.");
		}
	}
}
