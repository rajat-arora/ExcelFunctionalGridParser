import java.util.List;
import java.awt.Color;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * This class handles IO of the excel file Depends on Apache POI library
 * 
 * @author rajaarora
 *
 */
public class IOHandler {

	/**
	 * Constructor
	 * 
	 * @param none
	 * @return none
	 */
	public IOHandler() {
	}

	/**
	 * Returns the value of the current cell
	 * 
	 * @param cell
	 * @return String(Cell)
	 */
	private Object getCellValue(Cell cell) {
		// get the cell type
		// when we detect something as string store, otherwise null
		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_STRING:
			return cell.getStringCellValue();
		}
		return null;
	}

	/**
	 * This reads the excel file
	 * 
	 * @param excelFilePath
	 * @return List(Test Cases)
	 * @throws IOException
	 */
	public List<TestCases> readBooksFromExcelFile(String excelFilePath) throws IOException {

		// create new list to store columns in
		List<TestCases> list = new ArrayList<>();
		// take in path of excel file
		FileInputStream inputStream = new FileInputStream(new File(excelFilePath));

		// create new XSSF workBook,Sheet, style, etc..
		// set the sheet to 0, so we are at AFW tab
		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
		XSSFSheet firstSheet = workbook.getSheetAt(0);
		XSSFCellStyle style = workbook.createCellStyle();

		// Iterate through each row
		Iterator<Row> iterator = firstSheet.iterator();
		while (iterator.hasNext()) {
			// if we have a Row iterate through each cell
			XSSFRow nextRow = (XSSFRow) iterator.next();
			Iterator<Cell> cellIterator = nextRow.cellIterator();
			TestCases test = new TestCases(); // store the row (and it's
												// columns)

			// while we have a cell in our row
			while (cellIterator.hasNext()) {
				XSSFCell nextCell = (XSSFCell) cellIterator.next();
				// get the index of the cell we are on
				int columnIndex = nextCell.getColumnIndex();
				// depending on cell we are on. Set the value/color obtained
				// into test.
				switch (columnIndex) {
				case 0:
					test.setTag((String) getCellValue(nextCell));
					// System.out.println( (String) getCellValue(nextCell));
					break;
				case 1:
					test.setMajorFunctionalArea((String) getCellValue(nextCell));

					break;
				case 2:
					test.setGoogleDomainType((String) getCellValue(nextCell));

					break;
				case 3:
					test.setActivationType((String) getCellValue(nextCell));

					break;
				case 4:
					test.setAfWFeatures((String) getCellValue(nextCell));

					break;
				case 5:
					test.setPriority((String) getCellValue(nextCell));

					break;
				case 6:
					test.setAssign((String) getCellValue(nextCell));
					// set background color. Then create a new style. And
					// convert to hex
					test.setbColor(nextCell.getCellStyle().getFillForegroundColor());
					style = nextCell.getCellStyle();
					test.setHexColor(test.ColorToHex(style.getFillForegroundColorColor()));
					break;
				case 7:
					test.setAdd((String) getCellValue(nextCell));
					test.setbColor(nextCell.getCellStyle().getFillForegroundColor());
					style = nextCell.getCellStyle();
					test.setHexColor(test.ColorToHex(style.getFillForegroundColorColor()));
					break;
				case 8:
					test.setUpdate((String) getCellValue(nextCell));
					test.setbColor(nextCell.getCellStyle().getFillForegroundColor());
					style = nextCell.getCellStyle();
					test.setHexColor(test.ColorToHex(style.getFillForegroundColorColor()));
					;
					break;
				case 9:
					test.setDelete((String) getCellValue(nextCell));
					test.setbColor(nextCell.getCellStyle().getFillForegroundColor());
					style = nextCell.getCellStyle();
					test.setHexColor(test.ColorToHex(style.getFillForegroundColorColor()));
					break;
				case 10:
					test.setRemove((String) getCellValue(nextCell));
					test.setbColor(nextCell.getCellStyle().getFillForegroundColor());
					style = nextCell.getCellStyle();
					test.setHexColor(test.ColorToHex(style.getFillForegroundColorColor()));
					break;
				case 11:
					test.setTC((String) getCellValue(nextCell));

					break;
				case 12:
					test.setNumOfTC((String) getCellValue(nextCell));

					break;
				case 13:
					test.setRelease((String) getCellValue(nextCell));
					break;

				}

			}
			// test.notNull();

			// Add the excel row we just read (test) and store it into the main
			// list
			list.add(test);

		}

		// close the reading session and return what was read
		workbook.close();
		inputStream.close();
		// System.out.println(counter);
		return list;
	}

	/**
	 * This writes the data parsed into a new Excel File
	 * 
	 * @param List(test
	 *            cases)
	 * @param excelFilePath
	 */
	public void writeExcel(List<TestCases> l, String excelFilePath) {
		// create new workbook(buffer) and sheet
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet();
		// create new Headers for each row
		createHeaderRow(sheet);

		int rowCount = 0; // what row number we are on
		l.remove(0); // remove the headers from the list. We are already done
						// with them

		// For each row, write it's columns to the excel file(buffer)
		for (TestCases t : l) {
			XSSFRow row = sheet.createRow(++rowCount);
			writeTC(t, row, sheet);

		}

		// Set each columns to auto adjust it's size
		sheet.getWorkbook().getSheetAt(0).autoSizeColumn(0);
		sheet.getWorkbook().getSheetAt(0).autoSizeColumn(1);
		sheet.getWorkbook().getSheetAt(0).autoSizeColumn(2);
		sheet.getWorkbook().getSheetAt(0).autoSizeColumn(3);
		sheet.getWorkbook().getSheetAt(0).autoSizeColumn(4);
		sheet.getWorkbook().getSheetAt(0).autoSizeColumn(5);
		sheet.getWorkbook().getSheetAt(0).autoSizeColumn(6);
		sheet.getWorkbook().getSheetAt(0).autoSizeColumn(7);
		sheet.getWorkbook().getSheetAt(0).autoSizeColumn(8);
		sheet.getWorkbook().getSheetAt(0).autoSizeColumn(9);
		sheet.getWorkbook().getSheetAt(0).autoSizeColumn(10);
		sheet.getWorkbook().getSheetAt(0).autoSizeColumn(11);
		sheet.getWorkbook().getSheetAt(0).autoSizeColumn(12);
		sheet.getWorkbook().getSheetAt(0).autoSizeColumn(13);

		// Output buffer to excel file
		// If excel file fails writing or some excel file is that is named the
		// path is already open
		// throw Writing exception
		try (FileOutputStream outputStream = new FileOutputStream(excelFilePath)) {
			workbook.write(outputStream);
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	/**
	 * This creates the headers for the rows
	 * 
	 * @param XSSFSheet
	 */
	private void createHeaderRow(XSSFSheet sheet) {

		// create new cell style. Set it solid grey
		XSSFCellStyle cellStyle = sheet.getWorkbook().createCellStyle();
		cellStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
		cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);

		// create new font. Turn bold font on
		XSSFFont font = sheet.getWorkbook().createFont();
		font.setBold(true);
		cellStyle.setFont(font);
		// Create new row. Begin at row 0
		XSSFRow row = (XSSFRow) sheet.createRow(0);

		// create new Header
		// set it's style(color/font)
		// Set the name of the header
		XSSFCell tag = row.createCell(0);
		tag.setCellStyle(cellStyle);
		tag.setCellValue("Tag");

		XSSFCell majorFunctionalArea = row.createCell(1);
		majorFunctionalArea.setCellStyle(cellStyle);
		majorFunctionalArea.setCellValue("Major Functional Area");

		XSSFCell googleDomainType = row.createCell(2);
		googleDomainType.setCellStyle(cellStyle);
		googleDomainType.setCellValue("Google Domain Type");

		XSSFCell activationType = row.createCell(3);
		activationType.setCellStyle(cellStyle);
		activationType.setCellValue("Activation Type");

		XSSFCell AfWFeatures = row.createCell(4);
		AfWFeatures.setCellStyle(cellStyle);
		AfWFeatures.setCellValue("AfW Features");

		XSSFCell priority = row.createCell(5);
		priority.setCellStyle(cellStyle);
		priority.setCellValue("Priority");

		XSSFCell Assign = row.createCell(6);
		Assign.setCellStyle(cellStyle);
		Assign.setCellValue("Assign Pre Activation");

		XSSFCell Add = row.createCell(7);
		Add.setCellStyle(cellStyle);
		Add.setCellValue("Post Activation - Add");

		XSSFCell Update = row.createCell(8);
		Update.setCellStyle(cellStyle);
		Update.setCellValue("Post Activation - Update");

		XSSFCell Delete = row.createCell(9);
		Delete.setCellStyle(cellStyle);
		Delete.setCellValue("Post Activation - Delete");

		XSSFCell Remove = row.createCell(10);
		Remove.setCellStyle(cellStyle);
		Remove.setCellValue("Post Activation - Remove");

		XSSFCell TC = row.createCell(11);
		TC.setCellStyle(cellStyle);
		TC.setCellValue("MKS TestCase#");

		XSSFCell NumOfTC = row.createCell(12);
		NumOfTC.setCellStyle(cellStyle);
		NumOfTC.setCellValue("# of Manual TC");

		XSSFCell Release = row.createCell(13);
		Release.setCellStyle(cellStyle);
		Release.setCellValue("Release");

	}

	
	/**
	 * Writes each row to the excel file
	 * @param TestCases t
	 * @param XSSFRow
	 * @param XSSFSheet
	 */
	private void writeTC(TestCases t, XSSFRow row, XSSFSheet sheet) {

		//create new cell style 
		//Enable cell wraping
		XSSFCellStyle cellStyle = sheet.getWorkbook().createCellStyle();
		cellStyle.setWrapText(true);

		
		//create cell  @ row specified
		//set row value
		Cell cell = row.createCell(0);
		cell.setCellValue(t.getTag());

		
		//create cell @ row specified 
		//set row value
		//create new cell style
		//set wrap text to true
		//set the cell style
		cell = row.createCell(1);
		cell.setCellValue(t.getMajorFunctionalArea());
		cellStyle = sheet.getWorkbook().createCellStyle();
		cellStyle.setWrapText(true);
		cell.setCellStyle(cellStyle);

		cell = row.createCell(2);
		cell.setCellValue(t.getGoogleDomainType());
		cellStyle = sheet.getWorkbook().createCellStyle();
		cellStyle.setWrapText(true);
		cell.setCellStyle(cellStyle);

		cell = row.createCell(3);
		cell.setCellValue(t.getActivationType());
		cellStyle = sheet.getWorkbook().createCellStyle();
		cellStyle.setWrapText(true);
		cell.setCellStyle(cellStyle);

		cell = row.createCell(4);
		cell.setCellValue(t.getAfWFeatures());
		cellStyle = sheet.getWorkbook().createCellStyle();
		cellStyle.setWrapText(true);
		cell.setCellStyle(cellStyle);

		cell = row.createCell(5);
		cell.setCellValue(t.getPriority());
		cellStyle = sheet.getWorkbook().createCellStyle();
		cellStyle.setWrapText(true);
		cell.setCellStyle(cellStyle);

		
		//create cell @ row specified
		//set the value of the cell
		//set the color of the cell
		cell = row.createCell(6);
		cell.setCellValue(t.getAssign());
		setColor(cellStyle, sheet, t, cell, 0);

		// cellStyle.setFillForegroundColor(t.getbColor(1));
		cell = row.createCell(7);
		cell.setCellValue(t.getAdd());
		setColor(cellStyle, sheet, t, cell, 1);

		// cellStyle.setFillForegroundColor(t.getbColor(2));
		cell = row.createCell(8);
		cell.setCellValue(t.getUpdate());
		setColor(cellStyle, sheet, t, cell, 2);

		// cellStyle.setFillForegroundColor(t.getbColor(3));
		cell = row.createCell(9);
		cell.setCellValue(t.getDelete());
		setColor(cellStyle, sheet, t, cell, 3);

		// cellStyle.setFillForegroundColor(t.getbColor(4));
		cell = row.createCell(10);
		cell.setCellValue(t.getRemove());
		setColor(cellStyle, sheet, t, cell, 4);

		cell = row.createCell(11);
		cell.setCellValue(t.getTC());
		cellStyle = sheet.getWorkbook().createCellStyle();
		cellStyle.setWrapText(true);
		cell.setCellStyle(cellStyle);

		cell = row.createCell(12);
		cell.setCellValue(t.getNumOfTC());
		cellStyle = sheet.getWorkbook().createCellStyle();
		cellStyle.setWrapText(true);
		cell.setCellStyle(cellStyle);

		cell = row.createCell(13);
		cell.setCellValue(t.getRelease());
		cellStyle = sheet.getWorkbook().createCellStyle();
		cellStyle.setWrapText(true);
		cell.setCellStyle(cellStyle);

	}

	
	/**
	 * Sets the cell color
	 * @param XSSFCellStyle
	 * @param XSSFSheet
	 * @param TestCases
	 * @param Cell
	 * @param number
	 */
	private void setColor(XSSFCellStyle cellStyle, XSSFSheet sheet, TestCases t, Cell cell, int number) {
		//Create new XSSFColor
		XSSFColor myColor = null;
		// Set the current cell to wrap text. And to have solid color
		cellStyle = sheet.getWorkbook().createCellStyle();
		cellStyle.setWrapText(true);
		cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
		//If the color isn't null, then covert the Hex color back to XSSFColor
		//and then set the cell color.
		if (t.getHexColor(number) != null) {
			myColor = new XSSFColor(Color.decode("#" + t.getHexColor(number)));
			cellStyle.setFillForegroundColor(myColor);
			cell.setCellStyle(cellStyle);
		}
	}

}
