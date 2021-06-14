package github.qh.convert.domain.html2word;

import github.qh.convert.api.Html2Word;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.poifs.filesystem.DirectoryEntry;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.InputStream;
import java.io.OutputStream;

/**
 * @author qu.hao
 * @date 2021-04-22- 5:46 下午
 * @email quhao.mi@foxmail.com
 */
@Slf4j
public class ConvertByPoi implements Html2Word {

    private final InputStream inputStream;

    private final OutputStream outputStream;

    public ConvertByPoi(InputStream inputStream, OutputStream outputStream) {
        this.inputStream = inputStream;
        this.outputStream = outputStream;
    }

    @Override
    public void convert() {
        try (POIFSFileSystem poifsFileSystem = new POIFSFileSystem()) {
            //这里是必须要设置编码的，不然导出中文就会乱码。
            //将字节数组包装到流中
            DirectoryEntry directory = poifsFileSystem.getRoot();
            directory.createDocument("", inputStream);
            poifsFileSystem.writeFilesystem(outputStream);

            outputStream.close();
            inputStream.close();
        } catch (Exception e) {
            log.error("导出失败", e);
        }
    }
}
