package ui;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;

import java.io.File;
import java.awt.*;

/**
 * Class used to select the folder the SCADA Map
 * @author Hannah Layton
 */
public class MapSelectionUI extends JPanel {
	private static final long serialVersionUID = 1L;
	/** File chooser */
	JFileChooser chooser;
	/** Directory of selected folder */
	File fileDirectory;
	/** Path of selected folder */
	File filePath;
	/**
	 * Constructor for MapSelectionUI class, creates an instance of a MapSelectionUI object
	 */
	public MapSelectionUI(String choosertitle) {
		chooser = new JFileChooser(); 
		chooser.setCurrentDirectory(new java.io.File("."));
		chooser.setDialogTitle(choosertitle);
		chooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
		//
		// disable the "All files" option.
		//
		chooser.setAcceptAllFileFilterUsed(false);
		FileNameExtensionFilter filter = new FileNameExtensionFilter("XLSX files", "xlsx");
		chooser.setFileFilter(filter);
		//    
		if (chooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) { 
			fileDirectory = chooser.getCurrentDirectory();
			filePath = chooser.getSelectedFile();
		}
		else {
			System.exit(ABORT);
		}
	}
	/**
	 * Size of button window
	 * @return dimensions of window
	 */
	public Dimension getPreferredSize(){
		return new Dimension(300, 100);
	}
	/**
	 * Returns the directory of the selected file
	 * @return the directory of the selected file
	 */
	public File getFileDirectory() {
		return fileDirectory;
	}
	/**
	 * Returns the path of the selected file
	 * @return the path of the selected file
	 */
	public File getFilePath() {
		return filePath;
	}
}
