package github.qh.convert.domain.html2word;

import github.qh.convert.api.Html2Word;
import lombok.extern.slf4j.Slf4j;
import org.docx4j.convert.in.xhtml.XHTMLImporterImpl;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

/**
 * @author qu.hao
 * @date 2021-04-22- 5:46 下午
 * @email quhao.mi@foxmail.com
 */
@Slf4j
public class ConvertByDoc4j implements Html2Word {

    private final InputStream inputStream;

    private final OutputStream outputStream;

    public ConvertByDoc4j(InputStream inputStream,OutputStream outputStream) {
        this.inputStream = inputStream;
        this.outputStream = outputStream;
    }

    @Override
    public void convert() {
        try {
            WordprocessingMLPackage wordPackage = WordprocessingMLPackage.createPackage();
            XHTMLImporterImpl htmlImporter = new XHTMLImporterImpl(wordPackage);
            List<Object> convert = htmlImporter.convert(inputStream, "");
            wordPackage.getMainDocumentPart().getContent().addAll(convert);

            // Save it as a DOCX document on disc.
            wordPackage.save(outputStream);
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
