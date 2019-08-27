package azusa.cell.chizai_excel;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Hello world!
 *
 */
public class App {
	public static void main(String[] args) {
		System.out.println("test");

		// ワークブック
		XSSFWorkbook workBook = null;
		// シート
		XSSFSheet sheet = null;
		// 出力ファイル
		FileOutputStream outPutFile = null;
		// 出力ファイルパス
		String outPutFilePath = null;
		// 出力ファイル名
		String outPutFileName = null;

		// エクセルファイルの作成
		try {

			// ワークブックの作成
			workBook = new XSSFWorkbook();

			// シートの設定
			sheet = workBook.createSheet();
			workBook.setSheetName(0, "発注表");
			sheet = workBook.getSheet("発注表");

			// 初期行の作成
			XSSFRow row = sheet.createRow(2);

			// 「タイトル」のセルスタイル設定
			XSSFCellStyle titleCellStyle = workBook.createCellStyle();
			XSSFCell cell = row.createCell(3);
			XSSFFont titleFont = workBook.createFont();
			titleFont.setFontName("ＭＳ ゴシック");
			titleFont.setFontHeightInPoints((short) 20);
			titleFont.setUnderline(XSSFFont.U_SINGLE);
			titleCellStyle.setFont(titleFont);
			cell.setCellStyle(titleCellStyle);

			// セルに「タイトル」を設定
			cell.setCellValue("発注表");
			create(workBook, 25);

			// セルに「表のヘッダ」を設定
			row = sheet.createRow(3);
			cell = row.createCell(1);
			cell.setCellValue("受取日");

			// 「計算結果」のセルスタイル設定
			create(workBook, 25);
			// セルに「表のヘッダ」を設定

			// エクセルファイルを出力
			try {
				// ファイルパス・ファイル名の指定
				outPutFilePath = "C:\\Users\\komatsu\\Desktop\\";
				outPutFileName = "test.xlsx";

				// エクセルファイルを出力
				outPutFile = new FileOutputStream(
						outPutFilePath + outPutFileName);
				workBook.write(outPutFile);

				System.out.println(
						"「" + outPutFilePath + outPutFileName +
								"」を出力しました。");

			} catch (IOException e) {
				System.out.println(e.toString());
			}

		} catch (Exception e) {
			System.out.println(e.toString());
		}

	}

	private static void create(XSSFWorkbook workBook, int num) {
		XSSFCellStyle resultCellStyle = workBook.createCellStyle();
		XSSFFont resultFont = workBook.createFont();
		resultFont.setFontName("ＭＳ ゴシック");
		resultFont.setFontHeightInPoints((short) num);
		resultCellStyle.setFont(resultFont);
		resultCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		resultCellStyle.setAlignment(HorizontalAlignment.CENTER);
		resultCellStyle.setFillForegroundColor(IndexedColors.WHITE.index);
		resultCellStyle.setBorderTop(BorderStyle.MEDIUM);
		resultCellStyle.setBorderBottom(BorderStyle.MEDIUM);
		resultCellStyle.setBorderRight(BorderStyle.MEDIUM);
		resultCellStyle.setBorderLeft(BorderStyle.MEDIUM);
	}

}
