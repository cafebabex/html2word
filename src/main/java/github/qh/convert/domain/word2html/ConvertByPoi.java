package github.qh.convert.domain.word2html;

import fr.opensagres.poi.xwpf.converter.core.ImageManager;
import fr.opensagres.poi.xwpf.converter.xhtml.XHTMLConverter;
import fr.opensagres.poi.xwpf.converter.xhtml.XHTMLOptions;
import github.qh.convert.api.Word2Html;
import jdk.nashorn.internal.runtime.ParserException;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.PicturesManager;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import org.w3c.dom.Document;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.HashMap;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import static org.apache.commons.codec.CharEncoding.UTF_8;

/**
 * @author quhao
 * poi方式word转html
 */
@Slf4j
public class ConvertByPoi implements Word2Html {

    private static final String DOC = "doc";
    private static final String DOCX = "docx";

    private static final Pattern CSS_PATTERN = Pattern.compile("\\.(\\w+)\\{(?s)(.+?)\\}");

    @Override
    public String convert(String fileName, InputStream word) {
        String fileType = fileName.substring(fileName.lastIndexOf(".") + 1);
        try {
            if (DOC.equalsIgnoreCase(fileType)) {
                return docToHtml(word);
            } else if (DOCX.equalsIgnoreCase(fileType)) {
                return docxToHtml(word);
            } else {
                throw new IllegalArgumentException("文件名称不合法");
            }
        } catch (Exception e) {
            log.error("导出文件失败", e);
            throw new ParserException("转换出错：" + e.getMessage());
        }

    }


    /**
     * doc转换为html
     *
     * @param file word 文件
     * @return 阿里云文件地址
     */
    private String docToHtml(InputStream file) throws IOException, ParserConfigurationException, TransformerException {
        HWPFDocument wordDocument = new HWPFDocument(file);
        WordToHtmlConverter wordToHtmlConverter = null;
        wordToHtmlConverter = new WordToHtmlConverter(DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument());
        //你要处理图片的方式，是保存到本地还是上传到文件服务器
        wordToHtmlConverter.setPicturesManager(docImageManger());
        wordToHtmlConverter.processDocument(wordDocument);

        Document htmlDocument = wordToHtmlConverter.getDocument();
        DOMSource domSource = new DOMSource(htmlDocument);
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        StreamResult streamResult = new StreamResult(outputStream);


        TransformerFactory tf = TransformerFactory.newInstance();
        Transformer serializer = tf.newTransformer();
        serializer.setOutputProperty(OutputKeys.ENCODING, UTF_8);
        serializer.setOutputProperty(OutputKeys.INDENT, "yes");
        serializer.setOutputProperty(OutputKeys.METHOD, "html");
        serializer.transform(domSource, streamResult);

        return formatHtmlRowCss(outputStream.toString(UTF_8));
    }

    private PicturesManager docImageManger() {
        return (a, b, suggestedName, d, e) -> {
            //TODO 这里假装上传文件到自己的服务器，然后返回图片地址
            ByteArrayInputStream byteArrayInputStream = new ByteArrayInputStream(a);
            return "";
        };
    }

    /**
     * docx转换为html
     *
     * @param file word 文件
     * @return 阿里云文件地址
     */
    private String docxToHtml(InputStream file) {
        try (XWPFDocument document = new XWPFDocument(file);
             ByteArrayOutputStream out = new ByteArrayOutputStream();
             OutputStreamWriter outputStreamWriter = new OutputStreamWriter(out, StandardCharsets.UTF_8)) {

            XHTMLOptions options = XHTMLOptions.create();
            // 存放图片的文件夹
            options.setImageManager(new DocxImageManager());

            XHTMLConverter xhtmlConverter = (XHTMLConverter) XHTMLConverter.getInstance();
            xhtmlConverter.convert(document, outputStreamWriter, options);
            return formatHtmlRowCss(out.toString(UTF_8));

        } catch (Exception e) {
            e.printStackTrace();
        }
        return "";
    }


    private static class DocxImageManager extends ImageManager {

        /**
         * 本次我们采用将图片上传到文件服务器，所以需要一个map去记录没个原始图片的url和传到服务器上真是路径的映射
         * 因为是栈上变量，所以没必要使用 ConcurrentHashMap
         */
        private final HashMap<String, String> urlMap = new HashMap<>();

        //如果是往本地写图片，则用这个构造函数，指定图片存放的文件夹
        public DocxImageManager(File baseDir, String imageSubDir) {
            super(baseDir, imageSubDir);
        }

        //如果是向网络上传输图片，则用这个
        public DocxImageManager() {
            super(null, null);
        }

        /**
         * 此方法表示你要如何处理这个图片
         *
         * @param imagePath 图片的原始路径（在html标签中的路径）
         * @param imageData 图片字节
         * @throws IOException 读取异常
         */
        @Override
        public void extract(String imagePath, byte[] imageData) throws IOException {
            //TODO 这里要把图片传到自己服务器，然后把地址关联上
            urlMap.put(imagePath, imagePath);
        }


        /**
         * 此方法返回一个字符串，该字符串会直接写在html中img标签的src中
         *
         * @param uri 图片的原始路径（在html标签中的路径）
         * @return 返回你要替换成的图片路径
         */
        @Override
        public String resolve(String uri) {
            return urlMap.getOrDefault(uri, "");
        }
    }


    /**
     * 将html的class标签指向的样式直接转换成行内style样式，
     * 因为我们的前端编辑器很low，没法解析class样式，所以只能后端做喽
     *
     * @param styMap 所有的class样式
     * @param doc    当前html文本
     * @return 转换后的html
     */
    private static String replaceClass(Map<String, String> styMap, org.jsoup.nodes.Document doc) {
        String classString = "class";
        String styleString = "style";
        Elements anyClass = doc.getElementsByAttribute(classString);
        for (Element element : anyClass) {
            String aClass = element.attr(classString);
            String[] allClassStr = (aClass.split(" "));
            for (String classStr : allClassStr) {
                String style = element.attr(styleString);
                String s = styMap.getOrDefault(classStr, "");
                String attributeValue = style + s;
                if (StringUtils.isNoneBlank(attributeValue)) {
                    element.attr(styleString, attributeValue);
                }
            }
            element.removeAttr(classString);
        }
        return doc.toString();
    }

    /**
     * 将html的class标签指向的样式直接转换成行内style样式
     *
     * @param html html文本
     * @return 转换后的文本
     */
    private static String formatHtmlRowCss(String html) {
        org.jsoup.nodes.Document jsoup = Jsoup.parse(html);
        String allCss = jsoup.getElementsByTag("style").html();
        Map<String, String> styleMap = new HashMap<>();

        Matcher matcher = CSS_PATTERN.matcher(allCss);
        while (matcher.find()) {
            styleMap.put(matcher.group(1), matcher.group(2));
        }
        return replaceClass(styleMap, jsoup);
    }

}


