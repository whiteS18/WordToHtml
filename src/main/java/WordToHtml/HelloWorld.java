package WordToHtml; /**
 *
 */


import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.PicturesManager;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.hwpf.usermodel.Picture;
import org.apache.poi.hwpf.usermodel.PictureType;
import org.apache.poi.xwpf.converter.core.BasicURIResolver;
import org.apache.poi.xwpf.converter.core.FileImageExtractor;
import org.apache.poi.xwpf.converter.xhtml.XHTMLConverter;
import org.apache.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
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
* @Description: word转html，去除域代码
* @Param:
* @return:
* @Author: CW Song
* @Date: 2019/8/26
*/
public class HelloWorld {

    public static void main(String arge[]) {
        try {
            convert2Html("C:/demo/附件2：银联卡业务统计规范（业务口径）2019A版 (1).docx","C:/demo/附件2：银联卡业务统计规范（业务口径）2019A版 (1).html");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void writeFile(String content, String path) {
        FileOutputStream fos = null;
        BufferedWriter bw = null;
        try {
            File file = new File(path);
            fos = new FileOutputStream(file);
            bw = new BufferedWriter(new OutputStreamWriter(fos,"utf-8"));
            bw.write(content);
        } catch (FileNotFoundException fnfe) {
            fnfe.printStackTrace();
        } catch (IOException ioe) {
            ioe.printStackTrace();
        } finally {
            try {
                if (bw != null){
                    bw.close();
                }
                if (fos != null){
                    fos.close();
                }
            } catch (IOException ie) {
                ie.printStackTrace();
            }
        }
    }

    public static void convert2Html(String fileName, String outPutFile)
            throws TransformerException, IOException,
            ParserConfigurationException {
        if(fileName.endsWith(".docx")){

            String fileOutName = outPutFile;
            XWPFDocument document = new XWPFDocument(new FileInputStream(fileName));
            XHTMLOptions options = XHTMLOptions.create().indent(4);
            // 导出图片
            File imageFolder = new File("C:/demo/");
            options.setExtractor(new FileImageExtractor(imageFolder));
            // URI resolver  word的html中图片的目录路径
            options.URIResolver(new BasicURIResolver("C:/demo/"));
            File outFile = new File(fileOutName);
            outFile.getParentFile().mkdirs();
            OutputStream out = new FileOutputStream(outFile);
            XHTMLConverter.getInstance().convert(document, out, options);
            out.close();

        }
        if(fileName.endsWith(".doc")) {
            HWPFDocument wordDocument = new HWPFDocument(new FileInputStream(fileName));//WordToHtmlUtils.loadDoc(new FileInputStream(inputFile));
            WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(
                    DocumentBuilderFactory.newInstance().newDocumentBuilder()
                            .newDocument());
            wordToHtmlConverter.setPicturesManager(new PicturesManager() {
                @Override
                public String savePicture(byte[] content,
                                          PictureType pictureType, String suggestedName,
                                          float widthInches, float heightInches) {
                    return suggestedName;
                }
            });
            wordToHtmlConverter.processDocument(wordDocument);
            //save pictures
            List pics = wordDocument.getPicturesTable().getAllPictures();
            if (pics != null) {
                for (int i = 0; i < pics.size(); i++) {
                    Picture pic = (Picture) pics.get(i);
                    System.out.println();
                    try {
                        pic.writeImageContent(new FileOutputStream("C:/demo/"
                                + pic.suggestFullFileName()));
                    } catch (FileNotFoundException e) {
                        e.printStackTrace();
                    }
                }
            }
            Document htmlDocument = wordToHtmlConverter.getDocument();
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            DOMSource domSource = new DOMSource(htmlDocument);
            StreamResult streamResult = new StreamResult(out);

            TransformerFactory tf = TransformerFactory.newInstance();
            Transformer serializer = tf.newTransformer();
            //字符编码
            serializer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
            serializer.setOutputProperty(OutputKeys.INDENT, "yes");
            serializer.setOutputProperty(OutputKeys.METHOD, "html");
            serializer.transform(domSource, streamResult);
            out.close();
            String content = out.toString();
            //去除目录域代码
            int index = content.lastIndexOf(" TOC \\");
            System.out.println("index============"+index);
            String str = content.substring(0,index);
            int i = str.lastIndexOf("<span");
            String substring = str.substring(i, index);
            int i1 = content.indexOf("</span", index);
            String substring1 = content.substring(i, i1+7);
            content = content.replace(substring1,"");
            writeFile(content, outPutFile);
        }

    }
}




