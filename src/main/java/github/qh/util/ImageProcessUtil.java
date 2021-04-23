package github.qh.util;

import lombok.extern.slf4j.Slf4j;
import org.apache.http.client.fluent.Request;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.InputStream;
import java.math.BigDecimal;

/**
 * @author qu.hao
 * @date 2021-04-22- 5:50 下午
 * @email quhao.mi@foxmail.com
 * 处理图片大小工具类
 */
@Slf4j
public class ImageProcessUtil {

    private static final BigDecimal MAX_WORD_IMAGE_WIGHT = BigDecimal.valueOf(550);
    private static final BigDecimal MAX_WORD_IMAGE_HEIGHT = BigDecimal.valueOf(750);

    private ImageProcessUtil(){

    }

    /**
     * 处理word中过大的图片，调整到适配a4纸的页面大小，参数可以通过上面两个值调整
     * @param content 导出文本
     * @return 调整后的文本
     */
    public static String processImage(String content) {
        Document parse = Jsoup.parse(content);
        Elements elements = parse.select("img");
        elements.forEach(
                element -> {
                    String url = element.attr("src");
                    BufferedImage sourceImg = null;
                    try {
                        InputStream image = Request.Get(url).execute().returnContent().asStream();
                        sourceImg = ImageIO.read(image);
                        BigDecimal width = BigDecimal.valueOf(sourceImg.getWidth());
                        BigDecimal height = BigDecimal.valueOf(sourceImg.getHeight());
                        if(width.compareTo(MAX_WORD_IMAGE_WIGHT) > 0){
                            height = height.divide(width.divide(MAX_WORD_IMAGE_WIGHT,1,BigDecimal.ROUND_CEILING),BigDecimal.ROUND_CEILING);
                            width = MAX_WORD_IMAGE_WIGHT;

                            //style="width: 610px; height: 407px;" width="549" height="304"
                            setAttr(element, width, height);
                        }
                        if(height.compareTo(MAX_WORD_IMAGE_HEIGHT) > 0){
                            width = width.divide(height.divide(MAX_WORD_IMAGE_HEIGHT,1,BigDecimal.ROUND_CEILING),BigDecimal.ROUND_CEILING);
                            height = MAX_WORD_IMAGE_HEIGHT;

                            setAttr(element, width, height);
                        }
                    } catch (Exception e) {
                        log.error("导出word计算图片尺寸异常",e);
                    }
                    //此步是为标签增加'</img>'结尾符号，防止报错
                    element.appendText("");
                }
        );
        //去掉"html"字符串，方便后面拼接
        return parse.select("body").outerHtml();
    }

    private static void setAttr(Element element, BigDecimal width, BigDecimal height) {
        element.attr("style", String.format("width: %spx; height: %spx;", width, height));
        element.attr("height", String.valueOf(height));
        element.attr("width", String.valueOf(width));
    }
}
