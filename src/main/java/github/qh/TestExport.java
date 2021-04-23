package github.qh;

import github.qh.context.HtmlText;
import github.qh.convert.ConvertByIceBlue;
import github.qh.convert.ConvertByPoi;
import github.qh.convert.api.Html2Word;
import github.qh.util.ImageProcessUtil;
import lombok.extern.slf4j.Slf4j;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;

/**
 * @author qu.hao
 * @date 2021-04-08- 10:24 上午
 * @email quhao.mi@foxmail.com
 */
@Slf4j
public class TestExport {

    public static void main(String[] args) throws IOException {
        String context = HtmlText.WORD_HEAD + ImageProcessUtil.processImage(HtmlText.CONTEXT) + HtmlText.WORD_FOOT;

        //这里是必须要设置编码的，不然导出中文就会乱码。
        byte[] b = context.getBytes(StandardCharsets.UTF_8);
        //将字节数组包装到流中
        InputStream inputStream = new ByteArrayInputStream(b);


        long l = System.currentTimeMillis();
        OutputStream outputStream1 = Files.newOutputStream(Paths.get("/Users/quhao/poi_word.doc"));
        Html2Word poiConvert = new ConvertByPoi(inputStream, outputStream1);
        poiConvert.convert();
        log.info("poi方式导出用时:{}ms",System.currentTimeMillis() - l);

        l = System.currentTimeMillis();
        inputStream = new ByteArrayInputStream(b);
        OutputStream outputStream2 = Files.newOutputStream(Paths.get("/Users/quhao/doc4j_word.doc"));
        Html2Word doc4jConvert = new ConvertByPoi(inputStream, outputStream2);
        doc4jConvert.convert();
        log.info("doc4j方式导出用时:{}ms",System.currentTimeMillis() - l);

        OutputStream outputStream3 = Files.newOutputStream(Paths.get("/Users/quhao/blue_word.doc"));
        Html2Word blueConvert = new ConvertByIceBlue(context, outputStream3);
        l = System.currentTimeMillis();
        blueConvert.convert();
        log.info("ice-blue方式导出用时:{}ms",System.currentTimeMillis() - l);

    }
}
