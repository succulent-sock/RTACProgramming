package mapObjects;

import java.io.FileInputStream;
import java.util.LinkedList;
import java.util.Queue;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

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
			if (sheet.getSheetName().toLowerCase().contains("digital") && sheet.getSheetName().toLowerCase().contains("output")) {
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
				if (cell.getStringCellValue().toLowerCase().contains("address")) {
					this.dnpAddressColumn = cell.getColumnIndex();
					break;
				}
			}
			if (dnpAddressColumn >= 0) {
				break;
			}
		}
		if (dnpAddressColumn < 0) {
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
				if (cell.getStringCellValue().toLowerCase().contains("ied device")) {
					this.slaveIEDDeviceColumn = cell.getColumnIndex();
					break;
				}
			}
			if (slaveIEDDeviceColumn >= 0) {
				break;
			}
		}
		if (slaveIEDDeviceColumn < 0) {
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
				if (cell.getStringCellValue().toLowerCase().contains("ied wordbit") || cell.getStringCellValue().toLowerCase().contains("relay element")) {
					this.slaveIEDWordbitColumn = cell.getColumnIndex();
					break;
				}
			}
			if (slaveIEDWordbitColumn >= 0) {
				break;
			}
		}
		if (slaveIEDWordbitColumn < 0) {
			throw new IllegalArgumentException("Slave IED Wordbit column could not be found in SCADA Map.");
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
				if (cell.getStringCellValue().toLowerCase().contains("point nomenclature") || cell.getStringCellValue().equals("Point Nomenclature Description")) {
					this.descriptionColumn = cell.getColumnIndex();
					break;
				}
			}
			if (descriptionColumn >= 0) {
				break;
			}
		}
		if (descriptionColumn < 0) {
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
		while (currentSheet.getRow(rowCount).getCell(dnpAddressColumn) == null || (currentSheet.getRow(rowCount).getCell(dnpAddressColumn).getCellType() != CellType.NUMERIC)) {
			rowCount++;
		}
		// Loops through all entries in the SCADA map
		while (currentSheet.getRow(rowCount) != null) {
			if (currentSheet.getRow(rowCount).getCell(dnpAddressColumn) == null && rowCount >= currentSheet.getLastRowNum()) {
				break;
			}
			else if (currentSheet.getRow(rowCount).getCell(slaveIEDDeviceColumn) == null ) {
				rowCount++;
				continue;
			}
			else {
				// Check if the SCADA entry is valid
				if (!currentSheet.getRow(rowCount).getCell(slaveIEDDeviceColumn).getStringCellValue().equals("")) {
					double address;
					if (currentSheet.getRow(rowCount).getCell(dnpAddressColumn).getCellType() == CellType.STRING && currentSheet.getRow(rowCount).getCell(dnpAddressColumn).getStringCellValue().contains("-")) {
						address = Double.valueOf(currentSheet.getRow(rowCount).getCell(dnpAddressColumn).getStringCellValue().split("-")[currentSheet.getRow(rowCount).getCell(dnpAddressColumn).getStringCellValue().split("-").length - 1]);
					}
					else if ((currentSheet.getRow(rowCount).getCell(dnpAddressColumn).getNumericCellValue() == 0.0) && !scadaEntries.isEmpty()) {
						address = currentSheet.getRow(rowCount - 1).getCell(dnpAddressColumn).getNumericCellValue();
					}
					else {
						address = currentSheet.getRow(rowCount).getCell(dnpAddressColumn).getNumericCellValue();
					}
					String wordbit;
					if (currentSheet.getRow(rowCount).getCell(slaveIEDWordbitColumn).getCellType() == CellType.NUMERIC) {
						wordbit = String.valueOf(currentSheet.getRow(rowCount).getCell(slaveIEDWordbitColumn).getNumericCellValue());
					}
					else {
						wordbit = currentSheet.getRow(rowCount).getCell(slaveIEDWordbitColumn).getStringCellValue();
					}
					if (wordbit.contains(":")) {
						String[] wordbits = wordbit.split(":");
						SCADAEntry scadaEntry1 = new SCADAEntry(address, currentSheet.getRow(rowCount).getCell(slaveIEDDeviceColumn).getStringCellValue(), wordbits[0], currentSheet.getRow(rowCount).getCell(descriptionColumn).getStringCellValue());
						SCADAEntry scadaEntry2 = new SCADAEntry(address, currentSheet.getRow(rowCount).getCell(slaveIEDDeviceColumn).getStringCellValue(), wordbits[1], currentSheet.getRow(rowCount).getCell(descriptionColumn).getStringCellValue());
						scadaEntry1.setSlaveIEDDevice(scadaEntry1.getSlaveIEDDevice().replace("-", "_").replace("/", "_"));
						scadaEntry2.setSlaveIEDDevice(scadaEntry2.getSlaveIEDDevice().replace("-", "_").replace("/", "_"));
						scadaEntries.add(scadaEntry1);
						scadaEntries.add(scadaEntry2);
					}
					else {
						SCADAEntry scadaEntry = new SCADAEntry(address, currentSheet.getRow(rowCount).getCell(slaveIEDDeviceColumn).getStringCellValue(), wordbit, currentSheet.getRow(rowCount).getCell(descriptionColumn).getStringCellValue());
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
			throw new IllegalArgumentException("SCADA Map failed to close.");
		}
		this.scadaEntries = scadaEntries;
	}
}
