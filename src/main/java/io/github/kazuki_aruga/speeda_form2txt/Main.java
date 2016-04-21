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
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

/**
 * speeda-form2txtのメインクラス。
 * 
 * @author k-aruga
 */
public class Main {

	/**
	 * log.
	 */
	private static final Log log = LogFactory.getLog(Main.class);

	/**
	 * 処理開始。
	 * 
	 * @param args
	 *            引数。
	 * @throws IOException
	 *             入出力例外。
	 * @throws FileNotFoundException
	 *             ファイルが存在しない。
	 * @throws InvalidFormatException
	 *             Excelのフォーマットエラー。
	 * @throws EncryptedDocumentException
	 *             パスワードが掛かっている。
	 */
	public static void main(String[] args) throws IOException, EncryptedDocumentException, InvalidFormatException {

		if (args.length < 2) {

			usage();
			return;
		}

		final File inputDir = new File(args[0]);

		if (!inputDir.isDirectory()) {

			usage();
			return;
		}

		final File outputDir = new File(args[1]);

		if (outputDir.exists()) {

			if (!outputDir.isDirectory()) {

				usage();
				return;
			}

			if (0 < outputDir.listFiles().length) {

				usage();
				return;
			}

		} else {

			log.info("出力先ディレクトリを作成します。:" + outputDir.getAbsolutePath());

			if (!outputDir.mkdirs()) {

				log.error("出力先ディレクトリの作成に失敗しました。");

				usage();
				return;
			}
		}

		try {

			// 変換処理本体を実行する
			Form2TxtControler.convertAll(inputDir, outputDir);

			log.info("すべての処理が正常に終了しました。");

		} catch (EncryptedDocumentException | InvalidFormatException | IOException e) {

			log.error("エラーが発生しました。", e);

			throw e;
		}
	}

	/**
	 * 使用方法を表示する。
	 */
	private static void usage() {

		System.out.println("使用方法： speeda-form2txt <input-directory> <output-directory>");
		System.out.println("  <input-directory>: Excelファイルのあるディレクトリを指定します。");
		System.out.println("  <output-directory>: テキストファイルを出力するディレクトリを指定します。存在しないディレクトリまたは空のディレクトリを指定してください。");

		log.warn("引数の指定に誤りがあります。");
	}

}
