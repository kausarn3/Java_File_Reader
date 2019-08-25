package demo;

//Download imported files from poi apache.
import java.io.FileInputStream;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFRow.CellIterator;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class excel_reader {

	public static void main(String[] args) {
		try {
			//Specify file path
			FileInputStream file = new FileInputStream("D://lecture//eclipse//a.xlsx");
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			//Index of sheet.
			XSSFSheet sheet = workbook.getSheetAt(0);
			//Row iterator.
			Iterator<Row> rowIterator = sheet.iterator();
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();
				//Cell iterator.
				Iterator<Cell> cellIterator = row.cellIterator();
				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					//Types of data present in excel file.
					switch (cell.getCellType()) {
					case STRING:
						System.out.print(cell.getStringCellValue() + "\t");
						break;
					case NUMERIC:
						System.out.print(cell.getNumericCellValue() + "\t");
						break;
					case BOOLEAN:
						System.out.print(cell.getBooleanCellValue() + "\t");
						break;
					default:

					}
					

				}
				System.out.println("");
			}
		} catch (Exception e) {

		}

	}
}
