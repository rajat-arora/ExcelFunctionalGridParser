import java.util.ArrayList;

import java.util.List;

import java.util.Map.Entry;
import java.util.TreeMap;

/**
 * This class parses each column
 * 
 * @author rajaarora
 *
 */
public class TestCaseParser {

	private int tcLocation; // Store the location of the Test case in the main
							// excel file list
	private String tc; // store the MKS tc's
	private int numberOfTc; // store number of tc's

	/**
	 * Constructor
	 * 
	 */
	public TestCaseParser() {

	}

	/**
	 * Parses each MKS tc's. To figure out the # of TC's
	 * 
	 * @param List(TestCases)
	 * @return List(TestCases)
	 */
	public List<TestCases> tcConverter(List<TestCases> l) {
		List<TestCaseParser> split = new ArrayList<TestCaseParser>(); // stores
																		// all
																		// Mks
																		// test
																		// cases
		TestCaseParser b = new TestCaseParser(); // Stores single MKS TC
		TreeMap<String, Integer> tmap = new TreeMap<String, Integer>(); // Store
																		// non-duplicate
																		// test
																		// cases

		// for each row in the list(EXCEL FIle)
		for (int i = 1; i < l.size(); i++) {
			// if our TC row is not null or "
			if (l.get(i).getTC() != null && l.get(i).getTC() != "") {
				String[] array;
				// split the TC's into an array
				if (findNumOfInstance(l.get(i).getTC(),',') > 1) {
					array = l.get(i).getTC().split(",", 0);
				} else {
					array = new String[] { l.get(i).getTC() };
				}
				
				//System.out.println("");
				// for each TC in the array
				for (int j = array.length - 1; j > -1; j--) {
					// if any TC in the array is not "" or ?
					// then store the tc's individually and
					// set there location in the main list
					if (!array[j].trim().equals("") && !array[j].trim().equals("?")) {
						// System.out.println(array[j].trim());
						b = new TestCaseParser();
						b.setTc(array[j].trim());
						b.setTcLocation(i);
						split.add(b);
					}
					// System.out.println(Integer.parseInt(array[j].trim()) + "
					// == " + i);
				}

			}

		}

		// fast O(n) + 0(1) = 0(n) time
		// for each tc in the main list
		// add it to the treemap
		// treemap will check for duplicates. If any. Will replace the new tc
		// with the old one.
		for (int i = split.size() - 1; i >= 0; i--) {
			// tmap.set(split.get(i));
			tmap.put(split.get(i).getTc(), split.get(i).getTcLocation());

		}

		// for (TestCaseParser te : lump){
		// if(!te.getTc().toLowerCase().contains("s")){
		// Integer temp1 =
		// Integer.parseInt(l.get(te.getTcLocation()).getNumOfTC());
		// String temp = Integer.toString(temp1 +1) ;
		// l.get(te.getTcLocation()).setNumOfTC(temp);
		//
		// }
		// }

		int count = 0; // store total number of tc
		// for each tc in the tree map
		for (Entry<String, Integer> entry : tmap.entrySet()) {
			// System.out.printf("Key : %s and Value: %s %n", entry.getKey(),
			// entry.getValue());
			// System.out.println( l.get(entry.getValue()).getNumOfTC());

			// if any of the colors is not null and is all green and does not
			// contain "(s)" in it
			if (l.get(entry.getValue()).getHexColor(0) != null && l.get(entry.getValue()).getHexColor(1) != null
					&& l.get(entry.getValue()).getHexColor(3) != null
					&& l.get(entry.getValue()).getHexColor(4) != null) {
				if (l.get(entry.getValue()).getHexColor(0).equals("92D050")
						&& l.get(entry.getValue()).getHexColor(2).equals("92D050")
						&& l.get(entry.getValue()).getHexColor(2).equals("92D050")
						&& l.get(entry.getValue()).getHexColor(3).equals("92D050")
						&& l.get(entry.getValue()).getHexColor(4).equals("92D050")) {
					if (!entry.getKey().toLowerCase().contains("(s)")) {
						// For the TC. Get the number of TC. And inc by 1
						Integer temp1 = Integer.parseInt(l.get(entry.getValue()).getNumOfTC());
						String temp = Integer.toString(temp1 + 1);
						l.get(entry.getValue()).setNumOfTC(temp);
						count++;
					}

				}
			}

		}
		System.out.println("Total # of TC: " + count);

		return l;

	}

	/**
	 * Gets the location of the TC in the main list(Excel file)
	 * 
	 * @return tcLocation
	 */
	public int getTcLocation() {
		return tcLocation;
	}

	/**
	 * Sets the location of the TC in the main list(Excel file)
	 * 
	 * @param tcLocation
	 * @return tcLocation
	 */
	public void setTcLocation(int tcLocation) {
		this.tcLocation = tcLocation;
	}

	/**
	 * Gets the number of TC's
	 * 
	 * @return numberOfTC
	 */
	public int getNumberOfTc() {
		return numberOfTc;
	}

	/**
	 * Sets number of TC's
	 * 
	 * @param numberOfTc
	 */
	public void setNumberOfTc(int numberOfTc) {
		this.numberOfTc = numberOfTc;
	}

	/**
	 * gets the MKS TC
	 * 
	 * @return tc
	 */
	public String getTc() {
		return tc;
	}

	/**
	 * Sets the MKS TC
	 * 
	 * @param tc
	 */
	public void setTc(String tc) {
		this.tc = tc;
	}

    public static int findNumOfInstance(String word, int find)
    {
        int count = 0;
        int index = 0;
        
        while (true)
        {
               int foundAt = word.indexOf(find, index);
               if (foundAt != -1)
               {
                     count++;
                     index = foundAt + 1;
               }
               else
               {
                     break;
               }
        }
        
        return count;
 }

}
