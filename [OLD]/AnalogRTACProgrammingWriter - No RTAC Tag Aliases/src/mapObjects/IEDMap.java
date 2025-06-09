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
	/** Full device name listed on IED Map */
	private String fullDeviceName;
	/** Device name listed on IED Map */
	private String deviceName;
	/** Relay Element column in IED Map */
	private int wordbitColumn = -1;
	/** HMI Point Name column in IED Map */
	private int hmiPointNameColumn = -1;
	/** Point Address column in IED Map */
	private int pointAddressColumn = -1;
	/** Description column in IED Map */
	private int descriptionColumn = -1;
	/** RTAC mark column in IED Map */
	private int rtacMarkColumn = -1;
	/** SCADA mark column in IED Map */
	private int scadaMarkColumn = -1;
	/** Relay elements included in IED Map */
	private TreeMap<String,IEDMapEntry> analogPoints;
	
	/**
	 * Converts a data map file into an easily manipulatable Java object with helpful attributes
	 * @param iedName - data map file stream to read from to add helpful attributes to the Java objects
	 */
	public IEDMap(FileInputStream iedName) {
		openIEDMap(iedName);
		setDeviceName();
		setWordbitColumn();
		setHmiPointNameColumn();
		setPointAddressColumn();
		setDescriptionColumn();
		setRtacMarkColumn();
		setScadaMarkColumn();
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
		if (workbook.getSheet("Analog Points") == null) {
			throw new IllegalArgumentException("IED Map does not have Analog Points sheet.");
		}
		String deviceName = workbook.getSheet("Analog Points").getRow(2).getCell(3).getStringCellValue();
		if (deviceName != null && !deviceName.equals("")) {
			this.fullDeviceName = deviceName;
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
	 * Returns the full device game for the IED Map
	 * @return the full device game for the IED Map
	 */
	public String getFullDeviceName() {
		return fullDeviceName;
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
	 * Returns the column that the HMI point names in the data map are contained in
	 * @return the column that the HMI point names in the data map are contained in
	 */
	public int getHmiPointNameColumn() {
		return hmiPointNameColumn;
	}

	/**
	 * Finds the column that the HMI point names in the data map are contained in & saves it to the Java object
	 */
	public void setHmiPointNameColumn() {
		for (Row row : workbook.getSheet("Analog Points")) {
			for (Cell cell : row) {
				if (cell.getCellType() == CellType.NUMERIC) {
					continue;
				}
				else if (cell.getStringCellValue().toLowerCase().contains("hmi") && cell.getStringCellValue().toLowerCase().contains("point") && cell.getStringCellValue().toLowerCase().contains("name")) {
					this.hmiPointNameColumn = cell.getColumnIndex();
					break;
				}
			}
			if (hmiPointNameColumn >= 0) {
				break;
			}
		}
		if (hmiPointNameColumn < 0) {
			DialogBoxUI.infoBox("HMI Point Name column could not be found in IED Map.", "");
			throw new IllegalArgumentException("HMI Point Name column could not be found in IED Map.");
		}
	}

	/**
	 * Returns the column that the point addresses in the data map are contained in
	 * @return the column that the point addresses in the data map are contained in
	 */
	public int getPointAddressColumn() {
		return pointAddressColumn;
	}

	/**
	 * Finds the column that the point addresses in the data map are contained in & saves it to the Java object
	 */
	public void setPointAddressColumn() {
		for (Row row : workbook.getSheet("Analog Points")) {
			for (Cell cell : row) {
				if (cell.getCellType() == CellType.NUMERIC) {
					continue;
				}
				else if (cell.getStringCellValue().toLowerCase().contains("point") && cell.getStringCellValue().toLowerCase().contains("address")) {
					this.pointAddressColumn = cell.getColumnIndex();
					break;
				}
			}
			if (pointAddressColumn >= 0) {
				break;
			}
		}
		if (pointAddressColumn < 0) {
			DialogBoxUI.infoBox("Point Address column could not be found in IED Map.", "");
			throw new IllegalArgumentException("Point Address column could not be found in IED Map.");
		}
	}

	/**
	 * Returns the column that the descriptions in the data map are contained in
	 * @return the column that the descriptions in the data map are contained in
	 */
	public int getDescriptionColumn() {
		return descriptionColumn;
	}

	/**
	 * Finds the column that the descriptions in the data map are contained in & saves it to the Java object
	 */
	public void setDescriptionColumn() {
		for (Row row : workbook.getSheet("Analog Points")) {
			for (Cell cell : row) {
				if (cell.getCellType() == CellType.NUMERIC) {
					continue;
				}
				else if (cell.getStringCellValue().toLowerCase().contains("description")) {
					this.descriptionColumn = cell.getColumnIndex();
					break;
				}
			}
			if (descriptionColumn >= 0) {
				break;
			}
		}
		if (descriptionColumn < 0) {
			DialogBoxUI.infoBox("Description column could not be found in IED Map.", "");
			throw new IllegalArgumentException("Description column could not be found in IED Map.");
		}
	}

	/**
	 * Returns the RTAC column in the data map are contained in
	 * @return the RTAC column in the data map are contained in
	 */
	public int getRtacMarkColumn() {
		return rtacMarkColumn;
	}

	/**
	 * Finds the RTAC column in the data map are contained in & saves it to the Java object
	 */
	public void setRtacMarkColumn() {
		for (Row row : workbook.getSheet("Analog Points")) {
			for (Cell cell : row) {
				if (cell.getCellType() == CellType.NUMERIC) {
					continue;
				}
				else if (cell.getStringCellValue().equals("RTAC")) {
					this.rtacMarkColumn = cell.getColumnIndex();
					break;
				}
			}
			if (rtacMarkColumn >= 0) {
				break;
			}
		}
		if (rtacMarkColumn < 0) {
			DialogBoxUI.infoBox("RTAC mark column could not be found in IED Map.", "");
			throw new IllegalArgumentException("RTAC mark column could not be found in IED Map.");
		}
	}

	/**
	 * Returns the SCADA column in the data map are contained in
	 * @return the SCADA column in the data map are contained in
	 */
	public int getScadaMarkColumn() {
		return scadaMarkColumn;
	}

	/**
	 * Finds the SCADA column in the data map are contained in & saves it to the Java object
	 */
	public void setScadaMarkColumn() {
		for (Row row : workbook.getSheet("Analog Points")) {
			for (Cell cell : row) {
				if (cell.getCellType() == CellType.NUMERIC) {
					continue;
				}
				else if (cell.getStringCellValue().equals("SCADA")) {
					this.scadaMarkColumn = cell.getColumnIndex();
					break;
				}
			}
			if (scadaMarkColumn >= 0) {
				break;
			}
		}
		if (scadaMarkColumn < 0) {
			DialogBoxUI.infoBox("SCADA mark column could not be found in IED Map.", "");
			throw new IllegalArgumentException("SCADA mark column could not be found in IED Map.");
		}
	}

	/**
	 * Returns a tree of entries in the data map
	 * @return a tree of entries in the data map
	 */
	public TreeMap<String, IEDMapEntry> getAnalogPoints() {
		return analogPoints;
	}

	/**
	 * Adds entries in the data map to a tree map
	 */
	public void setAnalogPoints() {
		XSSFSheet currentSheet = workbook.getSheet("Analog Points");
		analogPoints = new TreeMap<String, IEDMapEntry>();
		int row = 5;
		// Checks if data map is empty of wordbits
		if (currentSheet != null && currentSheet.getRow(row) != null && currentSheet.getRow(row).getCell(wordbitColumn) != null) {
			// Saves wordbit in row
			String currentRelayElement = "";
			if (currentSheet.getRow(row).getCell(wordbitColumn).getCellType() == CellType.NUMERIC) {
				currentRelayElement = String.valueOf(currentSheet.getRow(row).getCell(wordbitColumn).getNumericCellValue());
			}
			else {
				currentRelayElement = currentSheet.getRow(row).getCell(wordbitColumn).getStringCellValue();
			}
			// Loops through the rest of the entries in the data map
			while (currentSheet.getRow(row) != null && currentSheet.getRow(row).getCell(wordbitColumn) != null && !currentRelayElement.equals("")) {
				boolean rtacMark = false;
				boolean scadaMark = false;
				if (currentSheet.getRow(row).getCell(rtacMarkColumn).getStringCellValue().equals("X")) {
					rtacMark = true;
				}
				if (currentSheet.getRow(row).getCell(scadaMarkColumn).getStringCellValue().equals("X")) {
					scadaMark = true;
				}
				IEDMapEntry currentEntry = new IEDMapEntry(fullDeviceName, currentRelayElement, currentSheet.getRow(row).getCell(hmiPointNameColumn).getStringCellValue(), "AI", currentSheet.getRow(row).getCell(pointAddressColumn).getStringCellValue(), currentSheet.getRow(row).getCell(descriptionColumn).getStringCellValue(), rtacMark, scadaMark);
				analogPoints.put(currentRelayElement, currentEntry);
				row++;
				// If this was the last entry in the data map, the loop can end
				if (row > currentSheet.getLastRowNum() || currentSheet.getRow(row) == null || currentSheet.getRow(row).getCell(wordbitColumn) == null) {
					break;
				}
				if (currentSheet.getRow(row).getCell(wordbitColumn).getCellType() == CellType.NUMERIC) {
					currentRelayElement = String.valueOf(currentSheet.getRow(row).getCell(wordbitColumn).getNumericCellValue());
				}
				else {
					currentRelayElement = currentSheet.getRow(row).getCell(wordbitColumn).getStringCellValue();
				}
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
