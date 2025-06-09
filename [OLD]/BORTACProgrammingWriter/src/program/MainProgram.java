/**
 * 
 */
package program;

import java.io.File;
import java.io.FileInputStream;
import java.io.PrintWriter;
import java.util.TreeMap;

import io.TXTWriterIO;
import mapObjects.SCADAMap;
import ui.DialogBoxUI;
import ui.FolderSelectionUI;
import ui.MapSelectionUI;
import mapObjects.IEDMap;

/**
 * Main class when running the program
 * @author Hannah Layton
 */
public class MainProgram {
	/**
	 * Main method that runs the program
	 * @param args - possible arguments
	 * @throws Exception - any error that occurs when the program runs
	 */
	public static void main(String[] args) throws Exception {
		// Opens the file selection window to select the SCADA Map
		MapSelectionUI scadaMapSelector = new MapSelectionUI("Please select your SCADA Map.");
		File scadaPath = scadaMapSelector.getFilePath();
		FileInputStream scadaStream = new FileInputStream(scadaPath);
		SCADAMap scadaMap = new SCADAMap(scadaStream);
		scadaStream.close();
		// Opens the file selection window to select the folder of data maps
		FolderSelectionUI iedMapSelector = new FolderSelectionUI("Please select your folder of IED Maps.");
		File iedMapFolderPath = iedMapSelector.getFilePath();
		// Generates the structured text file
		PrintWriter writer = new PrintWriter(scadaPath.getParentFile() + "\\RTAC Binary Output Structured Text.txt", "UTF-8");
		TreeMap<String, IEDMap> iedMaps = new TreeMap<String, IEDMap>();
		// Loops through each file in the folder of data maps
		for (File iedMapPath : iedMapFolderPath.listFiles()) {
			// Checks if the file is not an data map
			if (!iedMapPath.getName().toLowerCase().contains(".xlsx") || !iedMapPath.getName().contains("Data_Map")) {
				continue;
			}
			// Adds ied map to Tree Map
			FileInputStream iedStream = new FileInputStream(iedMapPath);
			IEDMap iedMap = new IEDMap(iedStream);
			String deviceName = iedMap.getDeviceName();
			iedMaps.put(deviceName, iedMap);
			iedStream.close();
		}
		TXTWriterIO.structuredTextWriter(writer, iedMaps, scadaMap);
		writer.close();
		// Opens the completion dialog box
		DialogBoxUI.infoBox("Writing Complete!", "");
	}
}
