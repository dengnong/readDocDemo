import org.apache.poi.POIXMLDocument;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.PicturesManager;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.hwpf.model.PicturesTable;
import org.apache.poi.hwpf.usermodel.Bookmarks;
import org.apache.poi.hwpf.usermodel.Picture;
import org.apache.poi.hwpf.usermodel.PictureType;
import org.apache.poi.xwpf.converter.core.FileImageExtractor;
import org.apache.poi.xwpf.converter.core.FileURIResolver;
import org.apache.poi.xwpf.converter.xhtml.XHTMLConverter;
import org.apache.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.xwpf.usermodel.*;
import org.w3c.dom.Document;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.*;
import java.util.List;

/**
 * Created by 54472 on 2018/1/25.
 */
public class POI {

    //获取doc里的文本内容
    public static String getText(String filePath) {
        String str = null;
        try {
            FileInputStream inputStream = new FileInputStream(filePath);
            WordExtractor wordExtractor = new WordExtractor(inputStream);
            str = wordExtractor.getText();
            inputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return str;
    }

    //按段落获取doc的内容
    public static void getParagraphTest(String filePath) {
        String[] str = null;
        try {
            FileInputStream inputStream = new FileInputStream(filePath);
            WordExtractor wordExtractor = new WordExtractor(inputStream);
            str = wordExtractor.getParagraphText();
            for(int i = 0; i < str.length; i++) {
                System.out.println(str[i]);
                System.out.println("================================================= " + i);
            }
            inputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void getDocPicture(String filePath) {
        try {
            FileInputStream inputStream = new FileInputStream(filePath);
            HWPFDocument document = new HWPFDocument(inputStream);
            Bookmarks bookmarks = document.getBookmarks();
            PicturesTable picturesTable = document.getPicturesTable();
            System.out.println("图片数量: " + picturesTable.getAllPictures().size());
            List<Picture> list = picturesTable.getAllPictures();
            System.out.println(list.get(0).getWidth() + " " + list.get(0).getHeight() + " " + list.get(0).getSize());
//            Map<String, String> reslut = new HashMap<>();
//            int count = bookmarks.getBookmarksCount();
//            for(int i = 0; i < count; i++) {
//                Range range = new Range(bookmarks.getBookmark(i).getStart(),
//                        bookmarks.getBookmark(i).getEnd(), document);
//                CharacterRun cr = range.getCharacterRun(0);
//                if(picturesTable.hasPicture(cr)) {
//                    Picture pic = picturesTable.extractPicture(cr, true);
//                    System.out.println(pic.getSize() + " " + pic.getContent());
//                } else {
//                    if(range.text().equals("")) {
//                        reslut.put(bookmarks.getBookmark(i).getName(), null);
//                    } else {
//                        reslut.put(bookmarks.getBookmark(i).getName(), "\"" + range.text() + "\"");
//                    }
//                }
//            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    //把docx转换成html,有表格显示问题，图片显示正常
    public static void docxOutHTML(String filePath) throws IOException, ParserConfigurationException {

        long startTime = System.currentTimeMillis();
        XWPFDocument document = new XWPFDocument(new FileInputStream(filePath));
        XHTMLOptions options = XHTMLOptions.create().indent(4);
        File imageFolder = new File("C:\\Users\\54472\\Desktop\\test");
        options.setExtractor(new FileImageExtractor(imageFolder));
        options.URIResolver(new FileURIResolver(imageFolder));
        File outFile = new File("C:\\Users\\54472\\Desktop\\test\\testDocxToHtml.html");
        outFile.getParentFile().mkdir();
        OutputStream out = new FileOutputStream(outFile);
        XHTMLConverter.getInstance().convert(document, out, options);

    }

    //把doc转换成html，表格显示正常，图片显示正常
    public static void docOutHtml(String filePath) throws IOException, ParserConfigurationException, TransformerException {
        long startTime = System.currentTimeMillis();
        HWPFDocument document = new HWPFDocument(new FileInputStream(filePath));
        WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(DocumentBuilderFactory
            .newInstance().newDocumentBuilder().newDocument());
        wordToHtmlConverter.setPicturesManager(new PicturesManager() {
            @Override
            public String savePicture(byte[] bytes, PictureType pictureType, String s, float v, float v1) {
                File imagePath = new File("C:/Users/54472/Desktop/test2/");
                if(!imagePath.exists()) {
                    imagePath.mkdirs();
                }
                File file = new File("C:/Users/54472/Desktop/test2/" + s);
                try {
                    OutputStream os = new FileOutputStream(file);
                    os.write(bytes);
                    os.close();

                } catch (Exception e) {
                }
                return "C:/Users/54472/Desktop/test2/" + s;
            }
        });
        wordToHtmlConverter.processDocument(document);
        Document htmlDocument = wordToHtmlConverter.getDocument();
        File htmlFile = new File("C:\\Users\\54472\\Desktop\\test2\\testDocToHtml.html");
        OutputStream outputStream = new FileOutputStream(htmlFile);
        DOMSource domSource = new DOMSource(htmlDocument);
        StreamResult streamResult = new StreamResult(outputStream);

        TransformerFactory factory = TransformerFactory.newInstance();
        Transformer serializer = factory.newTransformer();
        serializer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
        serializer.setOutputProperty(OutputKeys.INDENT, "yes");
        serializer.setOutputProperty(OutputKeys.METHOD, "html");
        serializer.transform(domSource, streamResult);

        outputStream.close();
    }

    //获取表格单元的内容，无图片
    public static void getTableText(String filePath) throws IOException {
        StringBuffer tableText = new StringBuffer();
        XWPFDocument document = new XWPFDocument(POIXMLDocument.openPackage(filePath));
        List<XWPFTable> allTable = document.getTables();

        for(XWPFTable xwpfTable : allTable) {
            //获取表格行
            List<XWPFTableRow> rows = xwpfTable.getRows();
            for(XWPFTableRow xwpfTableRow : rows) {
                //获取表格单元格数据
                List<XWPFTableCell> cells = xwpfTableRow.getTableCells();
                for(XWPFTableCell xwpfTableCell : cells) {
                    List<XWPFParagraph> paragraphs = xwpfTableCell.getParagraphs();
                    for(XWPFParagraph xwpfParagraph : paragraphs) {
                        List<XWPFRun> runs = xwpfParagraph.getRuns();
                        for(int i = 0; i < runs.size(); i++) {
                            XWPFRun run = runs.get(i);
                            tableText.append(run.toString() + " ");
                        }
                    }
                }

            }
        }
        System.out.println(tableText.toString());
    }
}
