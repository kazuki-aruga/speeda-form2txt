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

import org.junit.Test;

/**
 * Form2Txtクラスのテストケース。
 * 
 * @author k-aruga
 */
public class Form2TxtTest {

	/**
	 * Excel97-2003形式のテスト。
	 * 
	 * @throws Exception
	 *             エラーが発生。
	 */
	@Test
	public void xlsToTxt() throws Exception {

		final File in = new File(Form2TxtTest.class.getResource("form10k.xls").getPath());
		final File out = File.createTempFile("speeda-from2txt-xlsToTxt", ".txt");
		out.deleteOnExit();
		Form2Txt.convert(in, out);
	}

	/**
	 * Excel2007形式のテスト。
	 * 
	 * @throws Exception
	 *             エラーが発生。
	 */
	@Test
	public void xlsxToTxt() throws Exception {

		final File in = new File(Form2TxtTest.class.getResource("form10k.xlsx").getPath());
		final File out = File.createTempFile("speeda-from2txt-xlsxToTxt", ".txt");
		out.deleteOnExit();
		Form2Txt.convert(in, out);
	}

}
