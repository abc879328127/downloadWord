package com.threeclear.report.util;


import com.alibaba.fastjson.JSONArray;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STVerticalJc;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

/**
 * 生成word文档工具类
 * 为了代码的可读性使用了瀑式函数的写法,所有的函数都是void的但是会返回一个上下文对象,承前启后
 * @see org.apache.poi.POIDocument
 * @author X_L
 * @create 2019/2/22
 */
public class WordUtil {

    /*
        核心对象 用于生成word
     */
    private XWPFDocument doc = null;

    /*
        用于输出word到硬盘
     */
    private FileOutputStream out = null;

    /*
        word存放目录
     */
    private String docFilePath;

    /*
        word名字
     */
    private String docFileName;

    /*
        默认字体
     */
    private final String DEFAULTFONTFAMILY = "微软雅黑";

    /*
        字体默认大小
     */
    private final Integer BIGTITLEFONTSIZE = 16;

    /*
        文档标题颜色 red
     */
    private final String BIGTITLECOLOR = "FF0000";

    /*
        此工具类上下文对象,  ps:为了使用瀑式函数,所以得在其自身注入自身对象
     */
    private static WordUtil wordUtil = null;

    /*
        私有化无参构造,防止外界直接创建对象
     */
    private WordUtil(){}

    private WordUtil(String path, String docName){
        doc = new XWPFDocument();
        this.docFileName = docName;
        this.docFilePath = path;
    }

    /**
     * 私有化构造函数 返回当前类中的私有化对象,确保在使用瀑式函数的过程中使用的上下文对象是同一个
     * @author X_L
     * @create 2019/2/26
     */
    public static WordUtil builderWordUtil(String path,String docName){
        return wordUtil = new WordUtil(path,docName);
    }

    /**
     * 设置标题内容
     * @param title  文本
     * @return
     */
    public WordUtil createdBigTitle(String title){
        this.setParagraphStyle(title,this.BIGTITLECOLOR,this.BIGTITLEFONTSIZE,true,ParagraphAlignment.CENTER,this.DEFAULTFONTFAMILY);
        return this.wordUtil;
    }

    /**
     * 通用文本设置
     * @param text  文本
     * @param fontSize 字体大小
     * @param boldFlag  是否粗体
     * @param center   居中方式
     * @param indent    首行缩进
     * @return
     */
    public WordUtil createdText(String text,Integer fontSize,boolean boldFlag,ParagraphAlignment center,int indent){
        this.setParagraphStyle(text,null,fontSize,boldFlag,center,indent,this.DEFAULTFONTFAMILY);
        return this.wordUtil;
    }


    /**
     * 通用文本设置
     * @param text  文本
     * @param fontSize 字体大小
     * @param boldFlag  是否粗体
     * @param center   居中方式
     * @return
     */
    public WordUtil createdText(String text,Integer fontSize,boolean boldFlag,ParagraphAlignment center,String fontFamily){
        this.setParagraphStyle(text,null,fontSize,boldFlag,center,"".equals(fontFamily) ? this.DEFAULTFONTFAMILY : fontFamily);
        return this.wordUtil;
    }

    /**
     * 设置表格
     * @param tableData     表格数据
     * @param tableTitle    表头数组
     * @return
     */
    public WordUtil setTableText(List<JSONArray> tableData,String[] tableTitle){
        int tableWidth = 8200;//表格总宽度
        //设置表头
        int rows = tableData.size()+1;
        int cols = tableTitle.length;
        XWPFTable table = this.doc.createTable(rows,cols);
        XWPFHelperTable xwpfHelperTable = new XWPFHelperTable();
        //设置表格宽度
        xwpfHelperTable.setTableWidthAndHAlign(table,"8200",STJc.CENTER);
        //设置表格行高
        xwpfHelperTable.setTableHeight(table,10,STVerticalJc.CENTER);
        //设置表头
        XWPFTableRow row1 = table.getRow(0);
        for(int t=0;t<tableTitle.length;t++){
            String title = tableTitle[t];
            this.setTableTitle(row1,title,t);
        }
        for(int i=0;i<tableData.size();i++){
            JSONArray jsonArray = tableData.get(i);
            XWPFTableRow row = table.getRow(i + 1);
            row.setHeight(3);
            for(int j=0;j<jsonArray.size();j++){
                XWPFTableCell cell = row.getCell(j);
                if(i % 2 == 0){
                    cell.setColor("FCFCFD");
                }else{
                    cell.setColor("F7F7F9");
                }
                XWPFParagraph xwpfParagraph = cell.addParagraph();
                xwpfParagraph.setAlignment(ParagraphAlignment.CENTER);
                XWPFRun run = xwpfParagraph.createRun();
                run.setText((String.valueOf(jsonArray.get(j))));
                //删除单元格中的第一个段落
                cell.removeParagraph(0);
            }
        }
        return this.wordUtil;
    }

    /**
     * 设置表头
     * @param row1
     * @param title  数据
     */
    private void setTableTitle(XWPFTableRow row1, String title,int index) {
        XWPFTableCell cell = row1.getCell(index);
        cell.setColor("BFD9E1");
        XWPFParagraph xwpfParagraph = cell.addParagraph();
        xwpfParagraph.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun run = xwpfParagraph.createRun();
        run.setBold(true);
        run.setText(String.valueOf(title));
        //删除单元格中的第一个段落
        cell.removeParagraph(0);
    }

    /**
     * 添加横线图片
     * @param redLineImagePath      红色横杠图片存放路径
     * @author X_L
     * @create 2019/2/25
     */
    public WordUtil setRedLine(String redLineImagePath){
        XWPFParagraph paragraph = this.doc.createParagraph();
        paragraph.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun run = paragraph.createRun();
        try {
            FileInputStream in = new FileInputStream(new File(redLineImagePath));
            run.addPicture(in,XWPFDocument.PICTURE_TYPE_JPEG,"line", Units.toEMU(420),Units.toEMU(5));
        } catch (Exception e) {
            e.printStackTrace();
        }
        return this.wordUtil;
    }


    /**
     * 添加图片
     * @param imageName     图片名称
     * @param image         图片
     * @return
     */
    public WordUtil createdImage(String imageName,File image){
        XWPFParagraph paragraph = this.doc.createParagraph();
        paragraph.setAlignment(ParagraphAlignment.THAI_DISTRIBUTE);
        XWPFRun run = paragraph.createRun();
        try {
            FileInputStream in = new FileInputStream(image);
            run.addPicture(in,XWPFDocument.PICTURE_TYPE_PNG,imageName,Units.toEMU(450),Units.toEMU(200));
        } catch (Exception e) {
            e.printStackTrace();
        }

        return this.wordUtil;
    }

    public WordUtil createdImage2(String imageName,File image){
        XWPFParagraph paragraph = this.doc.createParagraph();
        paragraph.setAlignment(ParagraphAlignment.THAI_DISTRIBUTE);
        XWPFRun run = paragraph.createRun();
        try {
            FileInputStream in = new FileInputStream(image);
            run.addPicture(in,XWPFDocument.PICTURE_TYPE_PNG,imageName,Units.toEMU(450),Units.toEMU(350));
        } catch (Exception e) {
            e.printStackTrace();
        }

        return this.wordUtil;
    }


    /**
     * 设置单元与单元之间间隔
     * @param breakType     间隔方式
     * @return
     */
    public WordUtil setNewPage(BreakType breakType){
        XWPFParagraph paragraph = this.doc.createParagraph();
        paragraph.createRun().addBreak(breakType);
        return this.wordUtil;
    }



    /**
     * 设置段落样式
     * @param text
     * @param color
     * @param fontSize
     * @param boldFlag
     * @param center
     * @param indent
     */
    public void setParagraphStyle(String text,String color,Integer fontSize,boolean boldFlag,ParagraphAlignment center,int indent,String fontFamily){
        XWPFParagraph paragraph = this.doc.createParagraph();
        paragraph.setAlignment(center);
        paragraph.setIndentationFirstLine(indent);
        this.setRunStyle(paragraph,text,color,fontSize,boldFlag,fontFamily);
    }

    /**
     * 设置段落样式
     * @param text
     * @param color
     * @param fontSize
     * @param boldFlag
     * @param center
     */
    public void setParagraphStyle(String text,String color,Integer fontSize,boolean boldFlag,ParagraphAlignment center,String fontFamily){
        XWPFParagraph paragraph = this.doc.createParagraph();
        paragraph.setAlignment(center);
        this.setRunStyle(paragraph,text,color,fontSize,boldFlag,fontFamily);
    }

    /**
     * 设置run样式
     * @param paragraph
     * @param text
     * @param color
     * @param fontSize
     * @param boldFlag
     */
    private void setRunStyle(XWPFParagraph paragraph, String text, String color, Integer fontSize, boolean boldFlag,String fontFamily) {
        XWPFRun run = paragraph.createRun();
        run.setText(text);
        run.setFontSize(fontSize);
        run.setFontFamily(fontFamily);
        run.setColor(color);
        run.setBold(boldFlag);
    }

    /**
     * 将文档写入硬盘
     * @author X_L
     * @create 2019/2/26
     */
    public boolean closeDoc() {
        boolean res = true;
        try {
            this.out = new FileOutputStream(this.docFilePath + this.docFileName);
            this.doc.write(out);
        } catch (Exception e) {
            e.printStackTrace();
            res = false;
        } finally {
            if (out != null) {
                try {
                    out.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return res;
    }

    /**
     * 瀑式函数连接,增加可读性
     * @return
     */
    public WordUtil and(){
        return this.wordUtil;
    }
}
