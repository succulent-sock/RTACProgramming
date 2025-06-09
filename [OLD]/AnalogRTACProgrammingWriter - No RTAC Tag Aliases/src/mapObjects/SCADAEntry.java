package mapObjects;

/**
 * Object representing an entry in the SCADA map
 * @author Hannah Layton
 */
public class SCADAEntry {
	/** Analog DNP Address of entry */
	private double dnpAddress;
	/** Slave IED Device of entry */
	private String slaveIEDDevice;
	/** Slave IED Wordbit of entry */
	private String wordbit;
	/** Slave IED DNP index of entry */
	private double index;
	/** Description for SCADA entry */
	private String description;
	/** Scaling for SCADA entry */
	private double scaling;
	
	 /**
	  * Creates a new SCADAEntry object with attributes
	  * @param dnpAddress - DNP address for the entry in the SCADA programming
	  * @param slaveIEDDevice - the device of the entry in the SCADA map
	  * @param wordbit - wordbit of the entry in the SCADA map
	  * @param index - DNP index of the entry in the SCADA map
	  * @param description - description of the entry in SCADA map
	  * @param scaling - scale factor of the entry in SCADA map
	  */
	public SCADAEntry(double dnpAddress, String slaveIEDDevice, String wordbit, double index, String description, String scaling) {
		setDnpAddress(dnpAddress);
		setSlaveIEDDevice(slaveIEDDevice);
		setWordbit(wordbit);
		setIndex(index);
		setDescription(description);
		setScaling(scaling);
	}

	/**
	 * Returns the DNP address for the entry in the SCADA programming
	 * @return the DNP address for the entry in the SCADA programming
	 */
	public double getDnpAddress() {
		return dnpAddress;
	}

	/**
	 * Sets the DNP address for the entry in the SCADA programming
	 * @param dnpAddress - DNP address for the entry in the SCADA programming
	 */
	public void setDnpAddress(double dnpAddress) {
		this.dnpAddress = dnpAddress;
	}

	/**
	 * Returns the device of the entry in the SCADA map
	 * @return the device of the entry in the SCADA map
	 */
	public String getSlaveIEDDevice() {
		return slaveIEDDevice;
	}

	/**
	 * Sets the device of the entry in the SCADA map to an attribute of the Java object
	 * @param slaveIEDDevice - the device of the entry in the SCADA map
	 */
	public void setSlaveIEDDevice(String slaveIEDDevice) {
		if ((slaveIEDDevice.contains("87TA") || slaveIEDDevice.contains("74TA") || slaveIEDDevice.contains("51TA") || slaveIEDDevice.contains("90TA")) && Character.isDigit(slaveIEDDevice.charAt(0))) {
			char[] c = slaveIEDDevice.toCharArray();
			char temp = c[0];
			c[0] = c[1];
			c[1] = temp;
			this.slaveIEDDevice = new String(c);
		}
		else {
			this.slaveIEDDevice = slaveIEDDevice;
		}
		this.slaveIEDDevice = this.slaveIEDDevice.split(" ")[0];
	}

	/**
	 * Returns the wordbit of the entry in the SCADA map 
	 * @return the wordbit of the entry in the SCADA map 
	 */
	public String getWordbit() {
		return wordbit;
	}

	/**
	 * Sets the wordbit of the entry in the SCADA map to an attribute of the Java object
	 * @param wordbit - wordbit of the entry in the SCADA map
	 */
	public void setWordbit(String wordbit) {
		this.wordbit = wordbit;
	}

	/**
	 * Returns the DNP index of the entry in the SCADA map
	 * @return the DNP index of the entry in the SCADA map
	 */
	public double getIndex() {
		return index;
	}

	/**
	 * Sets the DNP index of the entry in the SCADA map to an attribute of the Java object
	 * @param slaveIEDDNP - DNP index of the entry in the SCADA map
	 */
	public void setIndex(double slaveIEDDNP) {
		this.index = slaveIEDDNP;
	}

	/**
	 * Returns the description of the scada entry
	 * @return the description of the scada entry
	 */
	public String getDescription() {
		return description;
	}

	/**
	 * Sets the description of the scada entry
	 * @param description - description to add
	 */
	public void setDescription(String description) {
		this.description = description;
	}

	/**
	 * Returns the scale factor of the SCADA map entry
	 * @return the scale factor of the SCADA map entry
	 */
	public double getScaling() {
		return scaling;
	}

	/**
	 * Sets the scale factor of the SCADA map entry
	 * @param scaling - scale factor of the SCADA map entry
	 */
	public void setScaling(String scaling) {
		Integer scaleFactor = Integer.valueOf(scaling.split("x")[scaling.split("x").length - 1]);
		this.scaling = scaleFactor;
	}
}
