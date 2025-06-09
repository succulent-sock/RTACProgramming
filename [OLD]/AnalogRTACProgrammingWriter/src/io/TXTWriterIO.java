/**
 * 
 */
package io;

import java.io.PrintWriter;
import java.util.Queue;
import java.util.TreeMap;

import mapObjects.IEDMap;
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
		while (!scadaEntries.isEmpty()) {
			SCADAEntry currentEntry = scadaEntries.remove();
			IEDMap currentIEDMap = iedMaps.get(currentEntry.getSlaveIEDDevice());
			// If Slave IED Device has a data map
			if (currentIEDMap != null) {
				TreeMap<String, String> analogPoints = currentIEDMap.getAnalogPoints();
				String rtacAlias = analogPoints.get(currentEntry.getWordbit());
				// If wordbit match is found and has a valid RTAC alias
				if (rtacAlias != null && !rtacAlias.equals("")) {
					// Add first half of line
					int currentDNPAddress = (int) currentEntry.getDnpAddress();
					if (currentDNPAddress < 10) {
						outputString.append("SCADA_DNP.AI_0000" + currentDNPAddress + " := " + rtacAlias + ";                    		SCADA_DNP.AI_0000"  + currentDNPAddress + ".instMag := " + rtacAlias + ".instMag");
					}
					else if (currentDNPAddress < 100) {
						outputString.append("SCADA_DNP.AI_000" + currentDNPAddress + " := " + rtacAlias + ";                    		SCADA_DNP.AI_000"  + currentDNPAddress + ".instMag := " + rtacAlias + ".instMag");
					}
					else if (currentDNPAddress < 1000) {
						outputString.append("SCADA_DNP.AI_00" + currentDNPAddress + " := " + rtacAlias + ";                    		SCADA_DNP.AI_00"  + currentDNPAddress + ".instMag := " + rtacAlias + ".instMag");
					}
					else if (currentDNPAddress < 10000) {
						outputString.append("SCADA_DNP.AI_0" + currentDNPAddress + " := " + rtacAlias + ";                    		SCADA_DNP.AI_0"  + currentDNPAddress + ".instMag := " + rtacAlias + ".instMag");
					}
					else {
						outputString.append("SCADA_DNP.AI_" + currentDNPAddress + " := " + rtacAlias + ";                    		SCADA_DNP.AI_"  + currentDNPAddress + ".instMag := " + rtacAlias + ".instMag");
					}
					// Add scaling factor
					if (currentEntry.getScaling() != 0) {
						outputString.append(" * " + currentEntry.getScaling());
					}
					outputString.append(";                    		// " + currentEntry.getDescription() + "\n");
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
