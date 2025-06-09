/**
 * 
 */
package io;

import java.io.PrintWriter;
import java.util.Map.Entry;
import java.util.Queue;
import java.util.TreeMap;

import mapObjects.IEDMap;
import mapObjects.IEDMapEntry;
import mapObjects.SCADAEntry;
import mapObjects.SCADAMap;

/**
 * Class responsible for writing the output text file
 * @author Hannah Layton
 */
public class TXTWriterIO {

	/**
	 * Writes the output text file
	 * @param txt - the file in which to put the output text
	 * @param iedMaps - TreeMap of ied maps
	 * @param scadaMap - SCADA map
	 */
	public static void structuredTextWriter(PrintWriter txt, TreeMap<String, IEDMap> iedMaps, SCADAMap scadaMap) {
		Queue<SCADAEntry> scadaEntries = scadaMap.getScadaEntries();
		StringBuilder outputString = new StringBuilder();
		// Loop through all SCADA Entries
		String currentDeviceName = "";
		boolean operTrip = true;
		while (!scadaEntries.isEmpty()) {
			SCADAEntry currentEntry = scadaEntries.remove();
			// Used to add line breaks between devices
			if (!currentDeviceName.equals(currentEntry.getSlaveIEDDevice())) {
				outputString.append("\n");
			}
			currentDeviceName = currentEntry.getSlaveIEDDevice();
			IEDMap currentIEDMap = iedMaps.get(currentEntry.getSlaveIEDDevice());
			// If Slave IED Device has a data map
			if (currentIEDMap != null) {
				TreeMap<String, IEDMapEntry> binaryOutputs = currentIEDMap.getBinaryOutputs();
				IEDMapEntry currentIEDEntry = binaryOutputs.get(currentEntry.getWordbit());
				if (currentIEDEntry == null) {
					for (Entry<String, IEDMapEntry> b : binaryOutputs.entrySet()) {
						String index = String.valueOf(Double.valueOf(b.getValue().getIndex()));
						if (currentEntry.getWordbit().equals(index)) {
							currentIEDEntry = b.getValue();
							break;
						}
					}
				}
				// If wordbit match is found and has a valid RTAC alias
				if (currentIEDEntry != null && !currentIEDEntry.getRtacPointName().equals("")) {
					String rtacPointName = currentIEDEntry.getRtacPointName();
					// Add first half of line
					outputString.append(rtacPointName);
					if (operTrip) {
						outputString.append(".operTrip    	 := ");
					}
					else {
						outputString.append(".operClose    	 := ");
					}
					// Add second half of line
					outputString.append("SCADA_DNP.BO_");
					int currentDNPAddress = (int) currentEntry.getDnpAddress();
					if (currentDNPAddress < 10) {
						outputString.append("0000").append(currentDNPAddress);
					}
					else if (currentDNPAddress < 100) {
						outputString.append("000").append(currentDNPAddress);
					}
					else if (currentDNPAddress < 1000) {
						outputString.append("00").append(currentDNPAddress);
					}
					else if (currentDNPAddress < 10000) {
						outputString.append("0").append(currentDNPAddress);
					}
					else {
						outputString.append(currentDNPAddress);
					}
					if (operTrip) {
						outputString.append(".operTrip");
						operTrip = false;
					}
					else {
						outputString.append(".operClose");
						operTrip = true;
					}
					outputString.append(";			// " + currentEntry.getDescription() + "\n");
				}
				else {
					outputString.append("NO RTAC ALIAS WAS FOUND FOR DNP ADDRESS: " + (int) currentEntry.getDnpAddress() + "\n");
				}
			}
			else {
				outputString.append("NO DATA MAP WAS FOUND FOR: " + currentEntry.getSlaveIEDDevice() + "\n");
			}
		}
		txt.print(outputString);
	}
}
