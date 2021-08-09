import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.File;
import java.io.FileOutputStream;

public class WordTool {
    public static void main(String[] args) throws Exception {
        writeWordFile();
    }

    public static void writeWordFile() throws Exception {
        XWPFDocument doc = new XWPFDocument();// 创建Word文件
        XWPFParagraph p = doc.createParagraph();// 新建段落
        p.setAlignment(ParagraphAlignment.CENTER);// 设置段落的对齐方式
        XWPFRun r = p.createRun();//创建标题
        r.setText("2020年元日大型活动情况分析");
        r.setBold(true);//设置为粗体
        r.setColor("000000");//设置颜色
        r.setFontSize(21); //设置字体大小
        r.addCarriageReturn();//回车换行
        XWPFParagraph p1 = doc.createParagraph();
        p1.setAlignment(ParagraphAlignment.BOTH);

        XWPFRun c1 = p1.createRun();
        c1.setText("一、12月31日晚上各地将举行各类活动");
        c1.setColor("000000");
        c1.setFontSize(12);
        c1.addCarriageReturn();

        String filePath = "E:/doc/";
        String file = "newYear.doc";

        FileOutputStream fileOutputStream = new FileOutputStream(new File(filePath + file));

        doc.write(fileOutputStream);
        fileOutputStream.flush();
        fileOutputStream.close();
    }

}