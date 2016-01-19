import java.util.ArrayList;
import java.util.List;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Class: TestCases
 * 
 * @author rajaarora Description: This class contains the columns of the excel
 *         file Usage: When file is read we can store each row we iterate
 *         through. As a object
 */
public class TestCases {
	/*
	 * Stores tag
	 */
	private String tag;
	/*
	 * Stores Major Functional Area
	 */
	private String majorFunctionalArea;
	/*
	 * Stores Google Domain Type
	 */
	private String googleDomainType;

	/*
	 * Stores Activation Type
	 */
	private String activationType;
	/*
	 * Stores AfW Features
	 */
	private String AfWFeatures;
	/*
	 * Stores Priority
	 */
	private String priority;
	/*
	 * Stores Assign Pre Activation
	 */
	private String Assign;
	/*
	 * Stores Post Activation - Add
	 */
	private String Add;
	/*
	 * Stores Post Activation - Update
	 */
	private String Update;
	/*
	 * Stores Post Activation - Delete
	 */
	private String Delete;
	/*
	 * Stores Post Activation - Remove
	 */
	private String Remove;
	/*
	 * Stores MKS TestCase#
	 */
	private String TC;
	/*
	 * Stores # of Manual TC
	 */
	private String NumOfTC;
	/*
	 * Stores Release
	 */
	private String Release;
	/*
	 * Stores the background color as a short(XSSF Color)
	 */
	private List<Short> bColor = new ArrayList<Short>();
	/*
	 * Stores hex color after converted from short(XSSF Color)
	 */
	private List<String> hexColor = new ArrayList<String>();

	/**
	 * Constructor
	 * 
	 * @param none
	 * @return none
	 */
	public TestCases() {
	}

	/**
	 * Gets the Major Functional Area
	 * 
	 * @param none
	 * @return majorFunctionalArea
	 */
	public String getMajorFunctionalArea() {
		return majorFunctionalArea;
	}

	/**
	 * Sets the Major Functional Area. If it's null, will default to ""
	 * 
	 * @param majorFunctionalArea
	 * @return none
	 * 
	 */
	public void setMajorFunctionalArea(String majorFunctionalArea) {
		if (this.majorFunctionalArea == null) {
			this.majorFunctionalArea = majorFunctionalArea;
		}
	}

	/**
	 * Gets the Tag
	 * 
	 * @param none
	 * @return tag
	 */
	public String getTag() {
		return tag;
	}

	/**
	 * Sets the Tag. If it's null, will default to ""
	 * 
	 * @param tag
	 * @return tag
	 */
	public void setTag(String tag) {
		if (this.tag == null) {
			this.tag = "";
		}
		this.tag = tag;
	}

	/**
	 * Gets the Google Domain Type @ returngoogleDomainType
	 */
	public String getGoogleDomainType() {
		return googleDomainType;
	}

	/**
	 * Sets the Google Domain Type. If it's null, will default to ""
	 * 
	 * @param googleDomainType
	 * @return googleDomainType
	 */
	public void setGoogleDomainType(String googleDomainType) {
		if (this.googleDomainType == null) {
			this.googleDomainType = "";
		}
		this.googleDomainType = googleDomainType;
	}

	/**
	 * Gets the Activation Type
	 * 
	 * @return activationType
	 */
	public String getActivationType() {
		return activationType;
	}

	/**
	 * Sets the Activation Type. If it's null, will default to ""
	 * 
	 * @param activationType
	 * @return activationType
	 */
	public void setActivationType(String activationType) {
		if (this.activationType == null) {
			this.activationType = "";

		}
		this.activationType = activationType;
	}

	/**
	 * Gets the AfW Features
	 * 
	 * @return AfWFeatures
	 */
	public String getAfWFeatures() {
		return AfWFeatures;
	}

	/**
	 * Sets the AfW Features. If it's null, will default to ""
	 * 
	 * @param afWFeatures
	 * @return afWFeatures
	 */
	public void setAfWFeatures(String afWFeatures) {
		if (this.AfWFeatures == null) {
			this.AfWFeatures = "";
		}

		this.AfWFeatures = afWFeatures;
	}

	/**
	 * Gets the Priority
	 * 
	 * @return priority
	 */
	public String getPriority() {
		return priority;
	}

	/**
	 * Sets the Priority. If it's null, will default to ""
	 * 
	 * @param priority
	 * @return priority
	 */
	public void setPriority(String priority) {
		if (this.priority == null) {
			this.priority = "";
		}
		this.priority = priority;
	}

	/**
	 * Gets the Assign Pre Activation
	 * 
	 * @return Assign
	 */
	public String getAssign() {
		return Assign;
	}

	/**
	 * Sets the Assign Pre Activation. If it's null, will default to ""
	 * 
	 * @param Assign
	 * @return Assign
	 */
	public void setAssign(String assign) {
		if (this.Assign == null) {
			this.Assign = "";
		}
		this.Assign = assign;
	}

	/**
	 * Gets the Post Activation - Add
	 * 
	 * @return Add
	 */
	public String getAdd() {
		return Add;
	}

	/**
	 * Sets the Post Activation - Add. If it's null, will default to ""
	 * 
	 * @param activationType
	 * @return activationType
	 */
	public void setAdd(String add) {
		if (this.Add == null) {
			this.Add = "";
		}
		this.Add = add;
	}

	/**
	 * Gets the Post Activation - Update
	 * 
	 * @return Update
	 */
	public String getUpdate() {
		return Update;
	}

	/**
	 * Sets the Post Activation - Update. If it's null, will default to ""
	 * 
	 * @param update
	 * @return update
	 */
	public void setUpdate(String update) {
		if (this.Update == null) {
			this.Update = "";
		}
		this.Update = update;
	}

	/**
	 * Gets the Post Activation - Delete
	 * 
	 * @return Delete
	 */
	public String getDelete() {
		return this.Delete;
	}

	/**
	 * Sets the Post Activation - Delete. If it's null, will default to ""
	 * 
	 * @param delete
	 * @return delete
	 */
	public void setDelete(String delete) {
		if (this.Delete == null) {
			this.Delete = "";
		}
		this.Delete = delete;
	}

	/**
	 * Gets the MKS TestCase#
	 * 
	 * @return TC
	 */
	public String getTC() {
		return TC;
	}

	/**
	 * Sets the MKS TestCase#. If it's null, will default to ""
	 * 
	 * @param TC
	 * @return TC
	 */
	public void setTC(String tC) {
		if (this.TC == null) {
			this.TC = "";
		}
		this.TC = tC;
	}

	/**
	 * Gets the Release
	 * 
	 * @return Release
	 */
	public String getRelease() {
		return Release;
	}

	/**
	 * Sets the Release. If it's null, will default to ""
	 * 
	 * @param Release
	 * @return Release
	 */
	public void setRelease(String release) {
		if (this.Release == null) {
			this.Release = "";
		}
		this.Release = release;
	}

	/**
	 * Gets # of Manual TC
	 * 
	 * @return NumOfTC
	 */
	public String getNumOfTC() {
		return NumOfTC;
	}

	/**
	 * Sets the # of Manual TC. If it's null, will default to ""
	 * 
	 * @param numOfTC
	 * @return numOfTC
	 */
	public void setNumOfTC(String numOfTC) {
		if (this.NumOfTC == null) {
			this.NumOfTC = "";
		}
		this.NumOfTC = numOfTC;
	}

	/**
	 * Converts all object to string
	 */
	public String toString() {
		return tag + "," + majorFunctionalArea + "," + googleDomainType + "," + activationType + "," + AfWFeatures + ","
				+ priority + "," + Assign + "(" + getHexColor(0) + ")" + "," + Add + "," + Update + "," + Delete + ","
				+ TC + "," + NumOfTC + "," + Release;
	}

	/**
	 * Gets the Remove
	 * 
	 * @return Remove
	 */
	public String getRemove() {
		return Remove;
	}

	/**
	 * Sets the Remove. If it's null, will default to ""
	 * 
	 * @param Remove
	 * @return Remove
	 */
	public void setRemove(String remove) {
		if (this.Remove == null) {
			this.Remove = "";
		}
		this.Remove = remove;
	}

	/**
	 * Gets the Background color
	 * 
	 * @return bColor at index i
	 */
	public short getbColor(int i) {
		return bColor.get(i);
	}

	/**
	 * Sets the Background color. If it's null, will default to ""
	 * 
	 * @param bColor
	 * @return bColor
	 */
	public void setbColor(short s) {
		this.bColor.add(s);
	}

	/**
	 * Converts XSSF Color to hexColor
	 * 
	 * @param XSSFColor
	 * @return hexColor
	 */
	public String ColorToHex(XSSFColor color) {
		if (color != null) {
			return ((XSSFColor) color).getARGBHex().substring(2, 8);
		}

		return null;

	}

	/**
	 * Gets the hexColor
	 * 
	 * @return hexColor at index i
	 */
	public String getHexColor(int i) {
		return hexColor.get(i);
	}

	/**
	 * Sets the hexColor. Adds it to the list of 4 colors for each row
	 * 
	 * @param hexColor
	 * @return hexColor
	 */
	public void setHexColor(String hexColor) {
		this.hexColor.add(hexColor);
	}

}
