import org.junit.Test;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by 54472 on 2018/1/29.
 */
public class TestTemplate {

    @Test
    public void testReplace() throws IOException {
        String inputPath = "C:\\Users\\54472\\Desktop\\template.docx";
        String outputPath = "C:\\Users\\54472\\Desktop\\temp.docx";
        Map<String, String> map = new HashMap<>();
        map.put("name", "小明");
        map.put("sex", "男");
        map.put("address", "软件园");
        map.put("phone", "8888888");

        List<String[]> list = new ArrayList<String[]>();
        list.add(new String[]{"111", "222", "333"});
        list.add(new String[]{"aaa", "bbb", "ccc"});
        list.add(new String[]{"@@@", "###", "$$$"});
        TemplateReplace.replace(inputPath, outputPath, map, list);
    }
}
