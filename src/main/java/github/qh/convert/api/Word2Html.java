package github.qh.convert.api;

import java.io.InputStream;

/**
 * @author quhao
 */
public interface Word2Html {

    /**
     * word转html
     * @param fileName file名称
     * @param word file
     * @return html-string
     */
    String convert(String fileName,InputStream word);
}
