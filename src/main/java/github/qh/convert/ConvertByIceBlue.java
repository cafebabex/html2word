package github.qh.convert;

import com.spire.doc.Document;
import com.spire.doc.FileFormat;
import com.spire.doc.Section;
import com.spire.doc.documents.MarginsF;
import github.qh.convert.api.Html2Word;
import lombok.extern.slf4j.Slf4j;

import java.io.IOException;
import java.io.OutputStream;

/**
 * @author qu.hao
 * @date 2021-04-22- 5:46 下午
 * @email quhao.mi@foxmail.com
 */
@Slf4j
public class ConvertByIceBlue implements Html2Word {

    private final String context;

    private final OutputStream outputStream;

    public ConvertByIceBlue(String context, OutputStream outputStream) {
        this.context = context;
        this.outputStream = outputStream;
    }

    @Override
    public void convert() {
        try {
            //新建Document对象
            Document document = new Document();
            //添加section
            Section sec = document.addSection();
            //设置页面边距
            MarginsF margins = sec.getPageSetup().getMargins();
            margins.setTop(50f);
            margins.setBottom(50f);
            //添加段落并写入HTML文本
            sec.addParagraph().appendHTML(context);
            //文档另存为doc
            document.saveToStream(outputStream, FileFormat.Doc);
        } catch (Exception e) {
            log.error("导出失败", e);
        }finally {
            try {
                outputStream.close();
            } catch (IOException e) {
                log.error("err",e);
            }
        }
    }
}
