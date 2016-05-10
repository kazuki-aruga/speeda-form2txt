/**
 * speeda-form2txt
 * 
 * Copyright (C) 2016 ARUGA, Kazuki.
 * 
 * Licensed under the Apache License, Version 2.0 (the &quot;License&quot;);
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an &quot;AS IS&quot; BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package io.github.kazuki_aruga.speeda_form2txt;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * SPEEDAからダウンロードしたExcelの有価証券報告書・半期報告書・四半期報告書の「事業の状況」をテキストファイルに出力する。
 * 
 * @author k-aruga
 */
public class Form2Txt {

	/**
	 * log.
	 */
	private static final Log log = LogFactory.getLog(Form2Txt.class);

	/**
	 * 事業の状況を記載したシート名。
	 */
	private static String[] sheetNames = { "事業の状況", "業績等の概要", "生産、受注及び販売の状況", "対処すべき課題", "経営上の重要な契約等", "研究開発活動" };

	/**
	 * SPEEDAからダウンロードしたExcelの有価証券報告書・半期報告書・四半期報告書の「事業の状況」をテキストファイルに出力する。
	 * 
	 * @param book
	 *            Excelファイル。
	 * @param text
	 *            テキストファイル。
	 * @throws IOException
	 *             入出力例外。
	 * @throws InvalidFormatException
	 *             Excelフォーマットエラー。
	 * @throws EncryptedDocumentException
	 *             パスワードが掛かっている。
	 */
	public static void convert(File book, File text)
			throws EncryptedDocumentException, InvalidFormatException, IOException {

		// ワークブックを開く
		try (Workbook workbook = WorkbookFactory.create(book)) {

			// 事業の状況シートが存在するか確認する
			final List<Sheet> sheetList = findSheets(workbook);
			if (sheetList.isEmpty()) {

				log.warn("<" + book.getAbsolutePath() + ">に[事業の状況]シートが存在しません。");
				return;
			}

			// 事業の状況シートの内容をテキストファイルに出力する
			try (PrintWriter writer = new PrintWriter(new FileWriter(text))) {

				for (Sheet sheet : sheetList) {

					writeAll(sheet, writer);
				}

				log.info("<" + book.getAbsolutePath() + ">の[事業の状況]を<" + text.getAbsolutePath() + ">に出力しました。");
			}
		}
	}

	/**
	 * ワークブックから事業の状況が記載されたシートを検索する。
	 * 
	 * @param workbook
	 *            ワークブック。
	 * @return シートの一覧。
	 */
	private static List<Sheet> findSheets(Workbook workbook) {

		final List<Sheet> result = new ArrayList<>();

		for (String sheetName : sheetNames) {

			// 事業の状況シートが存在するか確認する
			final int sheetIndex = workbook.getSheetIndex(sheetName);
			if (0 <= sheetIndex) {

				result.add(workbook.getSheetAt(sheetIndex));
			}
		}

		return result;
	}

	/**
	 * シートの内容を出力先に出力する。
	 * 
	 * @param sheet
	 *            シート。
	 * @param writer
	 *            出力先。
	 */
	private static void writeAll(Sheet sheet, PrintWriter writer) {

		for (Row row : sheet) {

			// 空行は飛ばす
			if (row.getFirstCellNum() != -1) {

				final Cell cell = row.getCell(0);
				// 出力すべきセルかどうか
				if (isPrintable(cell)) {

					final String outputValue = formatValue(cell.getStringCellValue());

					writer.println(outputValue);
				}
			}
		}
	}

	/**
	 * 出力すべきセルかどうかを判定する。
	 * 
	 * @param cell
	 *            セル。
	 * @return 出力すべき場合は<code>true</code>。
	 */
	private static boolean isPrintable(Cell cell) {

		// 文字列ではないセルは出力しない
		if (cell.getCellType() != Cell.CELL_TYPE_STRING) {

			return false;
		}

		// 罫線が引かれているセルは出力しない
		if (cell.getCellStyle().getBorderLeft() != CellStyle.BORDER_NONE) {

			return false;
		}

		final String value = cell.getStringCellValue();

		if (value == null) {

			return false;
		}

		return true;
	}

	/**
	 * 出力する文字列を整形する。
	 * 
	 * @param stringCellValue
	 *            セルの文字列。
	 * @return 整形済み文字列。
	 */
	private static String formatValue(String stringCellValue) {

		// nbspをスペースに変換する
		return stringCellValue.replace('\u00A0', ' ');
	}

}
