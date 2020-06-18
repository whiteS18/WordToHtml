package WordToHtml;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.PicturesManager;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.hwpf.usermodel.PictureType;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.xwpf.converter.core.BasicURIResolver;
import org.apache.poi.xwpf.converter.core.FileImageExtractor;
import org.apache.poi.xwpf.converter.xhtml.XHTMLConverter;
import org.apache.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.*;

/**
 * @description: word转html，去除目录域代码
 * @author: cw song
 * @create: 2019-08-14 10:51
 * @UpdateUser: cwsong
 * @update: 2019-08-14 10:51
 * @updateRemark:
 **/
public class WordToHtml {

    public static void main(String[] args) throws Exception {
        String fileName = Constant.FILE_NAME;
        String sourceFile = Constant.FILE_SOURCE_PATH;
        String htmlFile = Constant.FILE_HTML_PATH;
        String imgPath = Constant.IMG_PATH;
        String htmlPath = htmlFile + "/" + fileName.substring(0, fileName.lastIndexOf(".")) + ".html";
        docToHtml(imgPath, sourceFile+"\\"+fileName);
    }

    public static void docToHtml(final String imgPath, String sourceFileName) throws Exception {
        String content = null;
        FileOutputStream fos =null;
        BufferedWriter bw = null;

        File imgFile = new File(imgPath);
        if (!imgFile.exists()) {
            imgFile.mkdirs();
        }
        //doc为后缀的
        if(sourceFileName.endsWith(".doc")){
            HWPFDocument wordDocument = new HWPFDocument(new FileInputStream(sourceFileName));
            WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(
                    DocumentBuilderFactory.newInstance().newDocumentBuilder()
                            .newDocument());
            wordToHtmlConverter.setPicturesManager(new PicturesManager() {
                @Override
                public String savePicture(byte[] content, PictureType pictureType, String s, float v, float v1) {
                    File flie = new File(imgPath);
                    FileOutputStream fos = null;
                    try {
                        fos = new FileOutputStream(flie);
                        fos.write(content);
                        fos.close();
                    }catch (Exception e){

                    }
                    return imgPath+s;
                }
            });
            wordToHtmlConverter.processDocument(wordDocument);
            org.w3c.dom.Document htmlDocument = wordToHtmlConverter.getDocument();
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            DOMSource domSource = new DOMSource(htmlDocument);
            StreamResult streamResult = new StreamResult(baos);
            TransformerFactory tf = TransformerFactory.newInstance();
            Transformer serializer = tf.newTransformer();
            serializer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
            serializer.setOutputProperty(OutputKeys.INDENT, "yes");
            serializer.setOutputProperty(OutputKeys.METHOD, "html");
            serializer.transform(domSource, streamResult);
            baos.close();
            content = baos.toString();
            int index = content.lastIndexOf(" TOC \\");
            System.out.println("index============"+index);
            String str = content.substring(0,index);
            int i = str.lastIndexOf("<span");
            String substring = str.substring(i, index);
            int i1 = content.indexOf("</span", index);
            String substring1 = content.substring(i, i1+7);
            content = content.replace(substring1,"");
            System.out.println("str==================="+substring1);
        }//docx为后缀的
        else if(sourceFileName.endsWith(".docx")){
            FileInputStream fileInputStream = new FileInputStream(sourceFileName);
            //当你确定该文件可信任时使用该方法,避免大文件被误认为是zip炸弹,实际上不推荐上传大文件,会导致资源消耗过快
            ZipSecureFile.setMinInflateRatio(-1.0d);
            XWPFDocument document = new XWPFDocument(fileInputStream);
            XHTMLOptions options = XHTMLOptions.create().indent(4);
            options.setExtractor(new FileImageExtractor(new File(imgPath)));
            options.URIResolver(new BasicURIResolver(imgPath));
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            XHTMLConverter.getInstance().convert(document,baos,options);
            baos.close();
            content = baos.toString();
//            System.out.println(content);
        }
        else {
            System.out.println("请上传doc或者docx文件");
        }

        //jsoup转化并保存
      Document doc = Jsoup.parse(content);
        content=doc.html();
        try{
            File files = new File( Constant.FILE_HTML_PATH + "/" + Constant.FILE_NAME.substring(0, Constant.FILE_NAME.lastIndexOf(".")) + ".html");
            fos = new FileOutputStream(files);
            bw = new BufferedWriter(new OutputStreamWriter(fos,"UTF-8"));
            bw.write(content);
        }catch (Exception e){
            e.printStackTrace();
        }finally {
            if (bw!=null){
                bw.close();
            }
            if(fos!=null){
                fos.close();
            }
        }


    }


}

