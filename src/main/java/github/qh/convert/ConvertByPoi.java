package github.qh.convert;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.poifs.filesystem.DirectoryEntry;
import org.apache.poi.poifs.filesystem.DocumentEntry;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.ByteArrayInputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;

/**
 * @author qu.hao
 * @date 2021-04-22- 5:46 下午
 * @email quhao.mi@foxmail.com
 */
@Slf4j
public class ConvertByPoi {

    /**
     * 文本内容
     */
    private final String context;

    /**
     * 文件名称
     */
    private final String fileName;

    public ConvertByPoi(String context, String fileName) {
        this.context = context;
        this.fileName = fileName;
    }

    public void convert(OutputStream outputStream) {
        try (POIFSFileSystem poifsFileSystem = new POIFSFileSystem()) {
            //这里是必须要设置编码的，不然导出中文就会乱码。
            byte[] b = context.getBytes(StandardCharsets.UTF_8);
            //将字节数组包装到流中
            ByteArrayInputStream inputStream = new ByteArrayInputStream(b);
            DirectoryEntry directory = poifsFileSystem.getRoot();
            directory.createDocument(fileName, inputStream);
            poifsFileSystem.writeFilesystem(outputStream);

            inputStream.close();
            outputStream.close();
        } catch (Exception e) {
            log.error("导出失败", e);
        }
    }
}
