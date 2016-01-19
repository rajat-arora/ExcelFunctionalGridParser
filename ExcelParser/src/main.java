import java.io.IOException;
import java.util.List;


/**
 * Main Class invokes everything else. Takes no Arguments at the moment
 * @author rajaarora
 *
 */

public class main {

	public static void main(String[] args) {
		System.out.println("Starting...Please wait.."); 
		//create excel path and invoke IO class
		String excelFilePath = "test.xlsx";
		String outputExcel = "converted.xlsx";
		IOHandler reader = new IOHandler();
		
		List<TestCases> list = null; //stores all rows/columns of excel file

		
		System.out.println("Begining reading of file");
		
		
		//read the excel file and store all rows/columns
		//if wrong path or excel file. Throw IO exception 
		try {
			list = reader.readBooksFromExcelFile(excelFilePath);
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		//begin parsing 
		System.out.println("Finished reading. Now parsing");
		//for the excel file we just read. Initialize each #ofTestCases column to 0
		for (int i = 0; i < list.size(); i++) {
			list.get(i).setNumOfTC("0");
		}
		
		//invoke class TestCaseParser, pass it the excel file list
		TestCaseParser tp = new TestCaseParser();
		list = tp.tcConverter(list);
	
		
		
		
		//Write to file
		//pass the function writeExcel the list and name/location of what you want the output file
		System.out.println("Done parsing. Now writing");
		reader.writeExcel(list, "converted.xlsx");

		//Testing purposes (See output)
		for (int i = 0; i < list.size(); i++) {
			 //System.out.println(list.get(i).getHexColor(0));
			// System.out.println(list.get(i).toString());
			//System.out.println(list.get(i).getTC());

		}
		System.out.println("done");
		System.exit(0);

	}
}