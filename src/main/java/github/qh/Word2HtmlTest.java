package github.qh;

import github.qh.convert.api.Word2Html;
import github.qh.convert.domain.word2html.ConvertByPoi;

import java.io.FileInputStream;
import java.io.FileNotFoundException;

/**
 * @author quhao
 */
public class Word2HtmlTest {

    public static void main(String[] args) throws FileNotFoundException {
        Word2Html word2Html = new ConvertByPoi();

        String convert = word2Html.convert("aaa.doc", new FileInputStream("/Users/quhao/Downloads/word导入.doc"));
        System.out.println(convert);
    }
}
