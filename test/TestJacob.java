import org.junit.Test;

/**
 * Created by 54472 on 2018/1/29.
 */
public class TestJacob {

    @Test
    public void wordToHtml() {
        Jacob.DocToHtml(
                "C:\\Users\\54472\\Desktop\\卡片批量上传示例文档.docx",
                        "C:\\Users\\54472\\Desktop\\testHtml");
    }
}
