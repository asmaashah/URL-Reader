package ExcelReader;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Reader {
	static int total = 0;

	@SuppressWarnings("deprecation")
	public static void main(String[] args) throws IOException {
		try {
			FileInputStream file = new FileInputStream(new File("tweets.xlsx"));
			XSSFWorkbook tweets = new XSSFWorkbook(file);
			XSSFSheet sheet = tweets.getSheetAt(0);

			// open URLS excel document
			FileInputStream excelDoc = new FileInputStream(new File("URLS.xlsx"));
			XSSFWorkbook URLworkbook = new XSSFWorkbook(excelDoc);
			XSSFSheet URLSheet = URLworkbook.getSheetAt(0);

			Iterator<Row> rowIterator = sheet.iterator();

			for (int i = 0; i < 17411; i++) {
				Row row = rowIterator.next();
				Cell cell = row.getCell(7);

				String tweet = cell.getStringCellValue();

				// if the tweet has a website
				if (tweet.contains("http")) {
					total++;
					// Get a string with everything after the website
					String site = tweet.substring(tweet.indexOf("http") + 1);
					printURL(site, row.getRowNum(), URLSheet);

					// if there are multiple URLS
					while (site.contains("http")) {
						total++;
						site = site.substring(site.indexOf("http") + 1);
						printURL(site, row.getRowNum(), URLSheet);
					}
				}
			}

			file.close();
			URLworkbook.close();
			FileOutputStream output_file = new FileOutputStream("URLS.xlsx");
			URLworkbook.write(output_file);
			output_file.close();

		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
	}

	static public void printURL(String substring, int rowNum, XSSFSheet URLSheet) {
		// ignore everything after a space or newline
		String[] site2 = substring.split(" ");
		String[] siteURL = site2[0].split("\n");
		String finishedURL = "h" + siteURL[0];

		//enter URL
		Row row = URLSheet.createRow(total);
		Cell URLcell = row.createCell(0);
		URLcell.setCellValue((String) finishedURL);

		//enter original row in sheet
		Cell ORowCell = row.createCell(1);
		ORowCell.setCellValue((int) rowNum + 1);
	}
}
