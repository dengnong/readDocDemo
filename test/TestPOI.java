import org.junit.Test;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.TransformerException;
import java.io.IOException;

/**
 * Created by 54472 on 2018/1/25.
 */
public class TestPOI {
    @Test
    public void testReadTest() {
        System.out.println(POI.getText("C:\\Users\\54472\\Desktop\\wordTest.doc"));
    }

    @Test
    public void testReadTable() {
        POI.getParagraphTest("C:\\Users\\54472\\Desktop\\wordTest.doc");
    }

    @Test
    public void outHTML() throws IOException, ParserConfigurationException {
        POI.docxOutHTML("C:\\Users\\54472\\Desktop\\wordTest.docx");
    }

    @Test
    public void outHTML2() throws ParserConfigurationException, TransformerException, IOException {
        POI.docOutHtml("C:\\Users\\54472\\Desktop\\wordTest.doc");
    }

    @Test
    public void getTableText() throws IOException {
        POI.getTableText("C:\\Users\\54472\\Desktop\\wordTest.docx");
    }

    @Test
    public void getPic() throws IOException {
        POI.getDocPicture("C:\\Users\\54472\\Desktop\\wordTest.doc");
    }
}
