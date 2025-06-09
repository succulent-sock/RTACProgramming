/**
 * 
 */
package mapObjects;

import java.io.FileInputStream;
import java.util.LinkedList;
import java.util.Queue;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import ui.DialogBoxUI;

/**
 * A SCADA map file as an easily manipulatable Java object with helpful attributes
 * @author Hannah Layton
 */
public class SCADAMap {
	/** Variable used to manipulate an excel file, in this case a SCADA Map, with Java */
	private XSSFWorkbook workbook;
	/** Variable used to manipulate an excel sheet, in this case a SCADA Map, with Java */
	private Sheet currentSheet;
	/** Binary Output DNP Address column in SCADA Map */
	private int dnpAddressColumn = -1;
	/** Slave IED Device column in SCADA Map */
	private int slaveIEDDeviceColumn = -1;
	/** Slave IED Wordbit column in SCADA Map */
	private int slaveIEDWordbitColumn = -1;
	/** Slave IED DNP Index column in SCADA Map */
	private int slaveIEDDNPColumn = -1;
	/** Description column in SCADA Map */
	private int descriptionColumn = -1;
	/** List of entries included in SCADA Map */
	private Queue<SCADAEntry> scadaEntries;
	
	/**
	 * Converts a SCADA map file into an easily manipulatable Java object with helpful attributes
	 * @param scadaName - SCADA map file stream to read from to add helpful attributes to the Java objects
	 */
	public SCADAMap(FileInputStream scadaName) {
		openSCADAMap(scadaName);
		setCurrentSheet();
		setDnpAddressColumn();
		setSlaveIEDDeviceColumn();
		setSlaveIEDWordbitColumn();
		setSlaveIEDDNPColumn();
		setDescriptionColumn();
		setScadaEntries();
	}

	/**
	 * Returns a workbook object representing the SCADA map 
	 * @return a workbook object representing the SCADA map
	 */
	public XSSFWorkbook getSCADAMap() {
		return workbook;
	}

	/**
	 * Using the SCADA map file stream, a workbook object is created to read the SCADA map as an Excel file in Java
	 * @param scadaName - SCADA map file stream to read from to add helpful attributes to the Java objects
	 */
	public void openSCADAMap(FileInputStream scadaName) {
		try {
			this.workbook = new XSSFWorkbook(scadaName);
		} catch (Exception e) {
			DialogBoxUI.infoBox("Could not open SCADA Map.", "");
			throw new IllegalArgumentException("Could not open SCADA Map.");
		}
	}

	/**
	 * Returns the sheet of the SCADA map that is being read from
	 * @return the sheet of the SCADA map that is being read from
	 */
	public Sheet getCurrentSheet() {
		return currentSheet;
	}

	/**
	 * Sets the sheet of the SCADA map that is being read from
	 */
	public void setCurrentSheet() {
		for (Sheet sheet : workbook) {
			if (sheet.getSheetName().toLowerCase().contains("binary") && sheet.getSheetName().toLowerCase().contains("output")) {
				this.currentSheet = sheet;
				break;
			}
		}
	}

	/**
	 * Returns the column that the dnp addresses in the SCADA map are contained in
	 * @return the column that the dnp addresses in the SCADA map are contained in
	 */
	public int getDnpAddressColumn() {
		return dnpAddressColumn;
	}

	/**
	 * Finds the column that the dnp addresses in the SCADA map are contained in & saves it to the Java object
	 */
	public void setDnpAddressColumn() {
		for (Row row : currentSheet) {
			for (Cell cell : row) {
				if (cell.getStringCellValue().toLowerCase().contains("binary") && cell.getStringCellValue().toLowerCase().contains("address") && cell.getStringCellValue().toLowerCase().contains("output")) {
					this.dnpAddressColumn = cell.getColumnIndex();
					break;
				}
			}
			if (dnpAddressColumn >= 0) {
				break;
			}
		}
		if (dnpAddressColumn < 0) {
			DialogBoxUI.infoBox("DNP Address column could not be found in SCADA Map.", "");
			throw new IllegalArgumentException("DNP Address column could not be found in SCADA Map.");
		}
	}

	/**
	 * Returns the column that the slave ieds in the SCADA map are contained in
	 * @return the column that the slave ieds in the SCADA map are contained in
	 */
	public int getSlaveIEDDeviceColumn() {
		return slaveIEDDeviceColumn;
	}

	/**
	 * Finds the column that the slave ieds in the SCADA map are contained in & saves it to the Java object
	 */
	private void setSlaveIEDDeviceColumn() {
		for (Row row : currentSheet) {
			for (Cell cell : row) {
				if (cell.getStringCellValue().equals("Slave IED Device")) {
					this.slaveIEDDeviceColumn = cell.getColumnIndex();
					break;
				}
			}
			if (slaveIEDDeviceColumn >= 0) {
				break;
			}
		}
		if (slaveIEDDeviceColumn < 0) {
			DialogBoxUI.infoBox("Slave IED Device column could not be found in SCADA Map.", "");
			throw new IllegalArgumentException("Slave IED Device column could not be found in SCADA Map.");
		}
	}

	/**
	 * Returns the column that the slave ied wordbits in the SCADA map are contained in
	 * @return the column that the slave ied wordbits in the SCADA map are contained in
	 */
	public int getSlaveIEDWordbitColumn() {
		return slaveIEDWordbitColumn;
	}

	/**
	 * Finds the column that the slave ied wordbits in the SCADA map are contained in & saves it to the Java object
	 */
	private void setSlaveIEDWordbitColumn() {
		for (Row row : currentSheet) {
			for (Cell cell : row) {
				if (cell.getStringCellValue().equals("Slave IED Wordbit") || cell.getStringCellValue().contains("Relay Element") || cell.getStringCellValue().contains("IED Wordbit")) {
					this.slaveIEDWordbitColumn = cell.getColumnIndex();
					break;
				}
			}
			if (slaveIEDWordbitColumn >= 0) {
				break;
			}
		}
		if (slaveIEDWordbitColumn < 0) {
			DialogBoxUI.infoBox("Slave IED Wordbit column could not be found in SCADA Map.", "");
			throw new IllegalArgumentException("Slave IED Wordbit column could not be found in SCADA Map.");
		}
	}

	/**
	 * Returns the column that the slave ied DNP indexes in the SCADA map are contained in
	 * @return the column that the slave ied DNP indexes in the SCADA map are contained in
	 */
	public int getSlaveIEDDNPColumn() {
		return slaveIEDDNPColumn;
	}

	/**
	 * Finds the column that the slave ied DNP indexes in the SCADA map are contained in & saves it to the Java object
	 */
	private void setSlaveIEDDNPColumn() {
		for (Row row : currentSheet) {
			for (Cell cell : row) {
				if ((cell.getStringCellValue().toLowerCase().contains("relay") || cell.getStringCellValue().toLowerCase().contains("ied")) && cell.getStringCellValue().toLowerCase().contains("dnp") && cell.getStringCellValue().toLowerCase().contains("index")) {
					this.slaveIEDDNPColumn = cell.getColumnIndex();
					break;
				}
			}
			if (slaveIEDDNPColumn >= 0) {
				break;
			}
		}
		if (slaveIEDDNPColumn < 0) {
			DialogBoxUI.infoBox("Relay DNP Index column could not be found in SCADA Map.", "");
			throw new IllegalArgumentException("Relay DNP Index column could not be found in SCADA Map.");
		}
	}

	/**
	 * Returns the column that the descriptions in the SCADA map are contained in
	 * @return the column that the descriptions in the SCADA map are contained in
	 */
	public int getDescriptionColumn() {
		return descriptionColumn;
	}

	/**
	 * Finds the column that the descriptions in the SCADA map are contained in & saves it to the Java object
	 */
	public void setDescriptionColumn() {
		for (Row row : currentSheet) {
			for (Cell cell : row) {
				if (cell.getStringCellValue().toLowerCase().contains("nomenclature")) {
					this.descriptionColumn = cell.getColumnIndex();
					break;
				}
			}
			if (descriptionColumn >= 0) {
				break;
			}
		}
		if (descriptionColumn < 0) {
			DialogBoxUI.infoBox("Description column could not be found in SCADA Map.", "");
			throw new IllegalArgumentException("Description column could not be found in SCADA Map.");
		}
	}

	/**
	 * Returns a list of all entries in SCADA Map
	 * @return a list of all entries in SCADA Map
	 */
	public Queue<SCADAEntry> getScadaEntries() {
		return scadaEntries;
	}

	/**
	 * Sets the list of all entries in SCADA Map
	 */
	private void setScadaEntries() {
		Queue<SCADAEntry> scadaEntries = new LinkedList<SCADAEntry>();
		int rowCount = 1;
		// Find the first row containing a SCADA map entry
		while (currentSheet.getRow(rowCount).getCell(slaveIEDDNPColumn) == null || !(currentSheet.getRow(rowCount).getCell(slaveIEDDNPColumn).getNumericCellValue() >= 0)) {
			rowCount++;
		}
		// Loops through all entries in the SCADA map
		while (currentSheet.getRow(rowCount) != null) {
			if (currentSheet.getRow(rowCount).getCell(slaveIEDDNPColumn) == null && rowCount >= currentSheet.getLastRowNum()) {
				break;
			}
			else if (currentSheet.getRow(rowCount).getCell(slaveIEDDNPColumn) == null ) {
				rowCount++;
				continue;
			}
			else {
				if (!currentSheet.getRow(rowCount).getCell(slaveIEDWordbitColumn).getStringCellValue().equals("") && !currentSheet.getRow(rowCount).getCell(slaveIEDDeviceColumn).getStringCellValue().equals("")) {
					double address;
					// Check if the SCADA entry is valid
					if (currentSheet.getRow(rowCount).getCell(dnpAddressColumn).getCellType() == CellType.STRING && currentSheet.getRow(rowCount).getCell(dnpAddressColumn).getStringCellValue().contains("-")) {
						address = Double.valueOf(currentSheet.getRow(rowCount).getCell(dnpAddressColumn).getStringCellValue().split("-")[currentSheet.getRow(rowCount).getCell(dnpAddressColumn).getStringCellValue().split("-").length - 1]);
					}
					else {
						address = currentSheet.getRow(rowCount).getCell(dnpAddressColumn).getNumericCellValue();
					}
					if (currentSheet.getRow(rowCount).getCell(slaveIEDWordbitColumn).getStringCellValue().contains(":") && (currentSheet.getRow(rowCount).getCell(slaveIEDDNPColumn).getCellType() == CellType.STRING) && currentSheet.getRow(rowCount).getCell(slaveIEDDNPColumn).getStringCellValue().contains(":")) {
						String[] wordbits = currentSheet.getRow(rowCount).getCell(slaveIEDWordbitColumn).getStringCellValue().split(":");
						String[] indices = currentSheet.getRow(rowCount).getCell(slaveIEDDNPColumn).getStringCellValue().split(":");
						SCADAEntry scadaEntry1 = new SCADAEntry(address, currentSheet.getRow(rowCount).getCell(slaveIEDDeviceColumn).getStringCellValue(), wordbits[0], Double.valueOf(indices[0]), currentSheet.getRow(rowCount).getCell(descriptionColumn).getStringCellValue());
						SCADAEntry scadaEntry2 = new SCADAEntry(address, currentSheet.getRow(rowCount).getCell(slaveIEDDeviceColumn).getStringCellValue(), wordbits[1], Double.valueOf(indices[1]), currentSheet.getRow(rowCount).getCell(descriptionColumn).getStringCellValue());
						scadaEntry1.setSlaveIEDDevice(scadaEntry1.getSlaveIEDDevice().replace("-", "_").replace("/", "_"));
						scadaEntry2.setSlaveIEDDevice(scadaEntry2.getSlaveIEDDevice().replace("-", "_").replace("/", "_"));
						scadaEntries.add(scadaEntry1);
						scadaEntries.add(scadaEntry2);
					}
					else {
						SCADAEntry scadaEntry = new SCADAEntry(address, currentSheet.getRow(rowCount).getCell(slaveIEDDeviceColumn).getStringCellValue(), currentSheet.getRow(rowCount).getCell(slaveIEDWordbitColumn).getStringCellValue(), currentSheet.getRow(rowCount).getCell(slaveIEDDNPColumn).getNumericCellValue(), currentSheet.getRow(rowCount).getCell(descriptionColumn).getStringCellValue());
						scadaEntry.setSlaveIEDDevice(scadaEntry.getSlaveIEDDevice().replace("-", "_").replace("/", "_"));
						scadaEntries.add(scadaEntry);
					}
				}
				rowCount++;
			}
		}
		try {
			workbook.close();
		} catch (Exception e) {
			DialogBoxUI.infoBox("SCADA Map failed to close.", "");
			throw new IllegalArgumentException("SCADA Map failed to close.");
		}
		this.scadaEntries = scadaEntries;
	}
}
