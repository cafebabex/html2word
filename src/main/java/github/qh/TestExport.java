package github.qh;

import github.qh.context.HtmlText;
import github.qh.convert.ConvertByPoi;
import github.qh.util.ImageProcessUtil;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Paths;

/**
 * @author qu.hao
 * @date 2021-04-08- 10:24 上午
 * @email quhao.mi@foxmail.com
 */
public class TestExport {

    /**
     * html导出word的时候，必须将头信息修改成这个，不然打开word时候是web视图，而不是word的那个文本视图
     */
    private static final String WORD_HEAD = "<!DOCTYPE html PUBLIC \"-//W3C//DTD XHTML 1.0 Transitional//EN\" \"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd\">\r\n" +
            "<html xmlns=\"http://www.w3.org/TR/REC-html40\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:w=\"urn:schemas-microsoft-com:office:word\" xmlns:m=\"http://schemas.microsoft.com/office/2004/12/omml\">" +
            "<head><meta name=\"ProgId\" content=\"Word.Document\" /><meta name=\"Generator\" content=\"Microsoft Word 12\" />" +
            "<meta name=\"Originator\" content=\"Microsoft Word 12\" /> " +
            "<!--[if gte mso 9]><xml><w:WordDocument><w:View>Print</w:View></w:WordDocument></xml><[endif]-->" +
            "</head>";

    private static final String WORD_FOOT = "</html>";

    public static void main(String[] args) throws IOException {
        String context = HtmlText.CONTEXT;
        context = ImageProcessUtil.processImage(context);
        context = WORD_HEAD + context + WORD_FOOT;
        ConvertByPoi convert = new ConvertByPoi(context, HtmlText.TITLE);
        OutputStream outputStream = Files.newOutputStream(Paths.get("/Users/quhao/poi_word.doc"));
        convert.convert(outputStream);
    }
}
