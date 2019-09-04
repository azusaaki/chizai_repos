package azusa.cell.chizai_excel;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

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
		  // TODO 自動生成されたメソッド・スタブ

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

            // 「表のヘッダ」のセルスタイル設定
            XSSFCellStyle headerCellStyle = workBook.createCellStyle();
            XSSFFont headerFont = workBook.createFont();
            headerFont.setFontName("ＭＳ ゴシック");
            headerFont.setFontHeightInPoints((short) 25);
            headerCellStyle.setFont(headerFont);
            // headerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            headerCellStyle.setAlignment(HorizontalAlignment.CENTER);
            // headerCellStyle
            // .setFillForegroundColor(IndexedColors.PALE_BLUE.index);
            headerCellStyle.setBorderTop(BorderStyle.MEDIUM);
            headerCellStyle.setBorderBottom(BorderStyle.MEDIUM);
            headerCellStyle.setBorderRight(BorderStyle.MEDIUM);
            headerCellStyle.setBorderLeft(BorderStyle.MEDIUM);

            // セルに「表のヘッダ」を設定
            row = sheet.createRow(3);
            cell = row.createCell(1);
            cell.setCellStyle(headerCellStyle);
            cell.setCellValue("受取日");

            cell = row.createCell(2);
            cell.setCellStyle(headerCellStyle);
            cell.setCellValue("食材");

            cell = row.createCell(3);
            cell.setCellStyle(headerCellStyle);
            cell.setCellValue("個数");

            List<Order> orders = makeOrder();

            // 「計算結果」のセルスタイル設定
            create(workBook, 25);
            // セルに「表のヘッダ」を設定

            // エクセルファイルを出力
            try {
                // ファイルパス・ファイル名の指定
                outPutFilePath = "C:\\Users\\komatsu\\Desktop\\";
                outPutFileName = "発注表.xlsx";

                // エクセルファイルを出力
                outPutFile = new FileOutputStream(
                        outPutFilePath + outPutFileName);
                workBook.write(outPutFile);

                System.out.println(
                        "「" + outPutFilePath + outPutFileName + "」を出力しました。");

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

    private static List<Order> makeOrder() {
        List<Order> order = new ArrayList<Order>();

        // 名前、年収、年齢、身長
        order.add(new Order("1", "にんじん", 40, ""));
        order.add(new Order("2", "きゅうり", 40, ""));
        order.add(new Order("3", "なす", 40, ""));
        order.add(new Order("4", "とりにく", 40, ""));
        order.add(new Order("5", "ぶたにく", 40, ""));
        order.add(new Order("6", "ぎゅうにく", 40, ""));

        return order;
    }

}

//// 横
// for (int i = 3, j = 0; i < 13; i++, j++) {
// cell = row.createCell(i);
// cell.setCellStyle(headerCellStyle);
// if (i == 3) {
// cell.setCellValue("");
// } else {
// cell.setCellValue(j);
// }
// }
// // 縦
// for (int i = 6, j = 1; i < 15; i++, j++) {
// row = sheet.createRow(i);
// cell = row.createCell(3);
// cell.setCellStyle(headerCellStyle);
// cell.setCellValue(j);
// }
