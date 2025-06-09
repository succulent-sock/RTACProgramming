package ui;

import javax.swing.*;

import java.io.File;
import java.awt.*;

/**
 * Class used to select the folder of IED Maps for comparison - UNUSED
 * @author Hannah Layton
 */
public class FolderSelectionUI extends JPanel {
	private static final long serialVersionUID = 1L;
	/** File chooser */
	JFileChooser chooser;
	/** Directory of selected folder */
	File fileDirectory;
	/** Path of selected folder */
	File filePath;
	/**
	 * Constructor for FolderSelectionUI class, creates an instance of a FolderSelectionUI object
	 */
	public FolderSelectionUI(String choosertitle) {
		chooser = new JFileChooser(); 
		chooser.setCurrentDirectory(new java.io.File("."));
		chooser.setDialogTitle(choosertitle);
		chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
		//
		// disable the "All files" option.
		//
		chooser.setAcceptAllFileFilterUsed(false);
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
