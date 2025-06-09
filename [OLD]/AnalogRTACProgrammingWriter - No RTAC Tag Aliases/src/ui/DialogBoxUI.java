package ui;

import javax.swing.JOptionPane;

/**
 * Dialog box to indicate the program has finished running
 * @author Hannah Layton
 */
public class DialogBoxUI
{
	/**
	 * Dialog box to indicate the program has finished running
	 */
    public static void infoBox(String infoMessage, String titleBar)
    {
        JOptionPane.showMessageDialog(null, infoMessage, "", JOptionPane.INFORMATION_MESSAGE);
    }
}