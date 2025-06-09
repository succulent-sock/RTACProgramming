package mapObjects;

/**
 * Object class that is representative of an entry in a data map
 * @author Hannah Layton
 */
public class IEDMapEntry {
	/** Device Name */
	private String deviceName;
	/** Relay Element word bit */
	private String wordbit;
	/** HMI Point Name */
	private String hmiPointName;
	/** Point Type */
	private String pointType;
	/** Point Address */
	private int index;
	/** Description */
	private String description;
	/** Marked for SCADA */
	private boolean markedForSCADA;
	/** RTAC Point Name */
	private String rtacPointName = "";
	
	/**
	 * Object that is representative of an entry in a data map
	 * @param deviceName - name of the device listed in the data map
	 * @param wordbit - wordbit of the data map entry
	 * @param hmiPointName - HMI point name of the data map entry
	 * @param pointType - the type of the data map point
	 * @param index - DNP index of the data map entry
	 * @param description - point description of the data map entry
	 * @param scadaMark - whether or not the data map entry is marked to be sent to SCADA
	 */
	public IEDMapEntry(String deviceName, String wordbit, String hmiPointName, String pointType, double index, String description, boolean scadaMark) {
		setDeviceName(deviceName);
		setWordbit(wordbit);
		setHmiPointName(hmiPointName);
		setPointType(pointType);
		setIndex(index);
		setDescription(description);
		setMarkedForSCADA(scadaMark);
		setRtacPointName();
	}

	/**
	 * Returns the name of the device listed in the data map
	 * @return the name of the device listed in the data map
	 */
	public String getDeviceName() {
		return deviceName;
	}

	/**
	 * Sets the name of the device listed in the data map
	 * @param deviceName - the name of the device listed in the data map without modifications
	 */
	public void setDeviceName(String deviceName) {
		if ((deviceName.contains("87TA") || deviceName.contains("74TA") || deviceName.contains("51TA") || deviceName.contains("90TA")) && Character.isDigit(deviceName.charAt(0))) {
			char[] c = deviceName.toCharArray();
			char temp = c[0];
			c[0] = c[1];
			c[1] = temp;
			this.deviceName = new String(c);
		}
		else {
			this.deviceName = deviceName;
		}
		this.deviceName = this.deviceName.split(" ")[0];
	}

	/**
	 * Returns the wordbit of the data map entry
	 * @return the wordbit of the data map entry
	 */
	public String getWordbit() {
		return wordbit;
	}

	/**
	 * Sets the wordbit of the data map entry
	 * @param wordbit - wordbit of the data map entry
	 */
	public void setWordbit(String wordbit) {
		this.wordbit = wordbit;
	}

	/**
	 * Returns the HMI point name of the data map entry
	 * @return the HMI point name of the data map entry
	 */
	public String getHmiPointName() {
		return hmiPointName;
	}

	/**
	 * Sets the HMI point name of the data map entry
	 * @param hmiPointName - HMI point name of the data map entry
	 */
	public void setHmiPointName(String hmiPointName) {
		this.hmiPointName = hmiPointName;
	}

	/**
	 * Returns the type of point for the data map entry
	 * @return the type of point for the data map entry
	 */
	public String getPointType() {
		return pointType;
	}

	/**
	 * Set the type of point for the data map entry
	 * @param pointType - type of point for the data map entry
	 */
	public void setPointType(String pointType) {
		this.pointType = pointType;
	}

	/**
	 * Returns the DNP index of the data map entry
	 * @return the DNP index of the data map entry
	 */
	public int getIndex() {
		return index;
	}

	/**
	 * Sets the DNP index of the data map entry
	 * @param index - DNP index of the data map entry
	 */
	public void setIndex(double index) {
		this.index = (int) index;
	}

	/**
	 * Returns the point description of the data map entry
	 * @return the point description of the data map entry
	 */
	public String getDescription() {
		return description;
	}

	/**
	 * Sets the point description of the data map entry
	 * @param description - point description of the data map entry
	 */
	public void setDescription(String description) {
		this.description = description;
	}

	/**
	 * Returns whether the data map entry is marked for SCADA
	 * @return whether the data map entry is marked for SCADA
	 */
	public boolean isMarkedForSCADA() {
		return markedForSCADA;
	}

	/**
	 * Sets whether the data map entry is marked for SCADA
	 * @param markedForSCADA - whether or not the data map entry is marked for SCADA
	 */
	public void setMarkedForSCADA(boolean markedForSCADA) {
		this.markedForSCADA = markedForSCADA;
	}

	/**
	 * Returns the formulated RTAC point name that is created using other attributes in the data map entry
	 * @return the formulated RTAC point name that is created using other attributes in the data map entry
	 */
	public String getRtacPointName() {
		return rtacPointName;
	}

	/**
	 * Creates the RTAC point name using other attributes in the data map entry
	 */
	public void setRtacPointName() {
		if (markedForSCADA) {
			StringBuilder pointName = new StringBuilder();
			// Adds first half of the RTAC point name including device name and point type (AI)
			pointName.append(getDeviceName()).append("_DNP.").append(pointType).append("_");
			if (getIndex() < 10) {
				pointName.append("0000").append(getIndex());
			}
			else if (getIndex() < 100) {
				pointName.append("000").append(getIndex());
			}
			else if (getIndex() < 1000) {
				pointName.append("00").append(getIndex());
			}
			else if (getIndex() < 10000) {
				pointName.append("0").append(getIndex());
			}
			else {
				pointName.append(getIndex());
			}
			pointName.append("_");
			// Adds the last half of the RTAC point name using the HMI point name
			String endOfPointName = getHmiPointName().split(getDeviceName().split("_")[getDeviceName().split("_").length - 1])[getHmiPointName().split(getDeviceName().split("_")[getDeviceName().split("_").length - 1]).length - 1];
			endOfPointName = endOfPointName.substring(1).replace(".", "_");
			pointName.append(endOfPointName);
			this.rtacPointName = pointName.toString();
		}
	}
}
