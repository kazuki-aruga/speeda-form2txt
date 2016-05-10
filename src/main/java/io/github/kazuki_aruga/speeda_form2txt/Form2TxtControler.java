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
import java.io.FileFilter;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

/**
 * ディレクトリに存在するすべてのExcelファイルにForm2Txtを適用する。
 * 
 * @author k-aruga
 */
public class Form2TxtControler {

	/**
	 * 入力ディレクトリ以下にあるすべてのExcelファイルについて、Form2Txt処理を行って作成したファイルを、
	 * 出力フォルダ以下の同じ階層に出力する。
	 * 
	 * @param inputDir
	 *            入力ディレクトリ。
	 * @param outputDir
	 *            出力ディレクトリ。
	 * @throws IOException
	 *             入出力例外。
	 * @throws InvalidFormatException
	 *             Excelのフォーマットエラー。
	 * @throws EncryptedDocumentException
	 *             パスワードが掛かっている。
	 */
	public static void convertAll(File inputDir, File outputDir)
			throws EncryptedDocumentException, InvalidFormatException, IOException {

		final File[] listFiles = inputDir.listFiles(new FileFilter() {

			/**
			 * ディレクトリとExcelファイルのみリスト化する。
			 */
			@Override
			public boolean accept(File pathname) {

				if (pathname.isDirectory()) {

					return true;
				}

				if (pathname.getName().endsWith(".xls") || pathname.getName().endsWith(".xlsx")) {

					return true;
				}

				return false;
			}
		});

		for (File dirOrExcel : listFiles) {

			if (dirOrExcel.isDirectory()) {

				final File newOutputDir = new File(outputDir, dirOrExcel.getName());
				newOutputDir.mkdir();

				// 再帰的に処理を適用する
				convertAll(dirOrExcel, newOutputDir);

			} else {

				final File newTxtFile = new File(outputDir,
						outputDir.getName() + '_' + dirOrExcel.getName().replaceAll("[.](xls|xlsx)$", ".txt"));
				Form2Txt.convert(dirOrExcel, newTxtFile);
			}
		}
	}

}
