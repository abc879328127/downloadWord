package com.threeclear.report.controller;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.threeclear.report.tool.mcConfig;
import com.threeclear.report.util.ExcelUtils;
import com.threeclear.report.util.WordUtil;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.util.StringUtils;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.ResponseBody;
import sun.misc.BASE64Decoder;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.net.URLEncoder;
import java.util.List;

@Controller
public class downloadWord {


    //获取配置文件中的下载路径
    private static String qcWordReportFolder = null;

    static{
        try {
            qcWordReportFolder = mcConfig.getMCValue("report.reportWordFolder");
        }catch (Exception e){
            e.printStackTrace();
        }
    }

    /**
     * 前台是以这种拼接的格式，这种根据word展示样式所拼接的。图片是前台转换成svg格式，前后台没有分离。但是个人感觉与前后台分离是一样的
     * reportData: {
     *            title: "",
     *            titletwo: "",
     *            time1: "",   //第几期
     *            time2: "",   //时间
     *            chartitle1: "",  //第一段标题
     *             chartitle1Form: {      //第一段表格
     *                 formt: [],
     *             },
     *
     *            chartitle2: "",  //第二段标题
     *            chartitle2c: "",  //----
     *            chartitle2Image: {       //第二段图片
     *                chart2: [],
     *            },
     *            chartitle3: "",  //第三段标题
     *            chartitle3Image: {       //第三段图片
     *                chart3: [],
     *            },
     *
     *            chartitle4: "",  //第四段标题
     *            chartitle4Image: {       //第四段图片
     *                chart4: [],
     *            },
     *
     *            chartitle5: "",  //第五段标题
     *            chartitle5Image: {       //第五段图片
     *                chart5: [],
     *            },
     *
     *            chartitle6: "",  //第六段标题
     *            chartitle6Image: {       //第六段图片
     *                chart6: [],
     *            },
     *            chartitle7: "",  //第七段标题
     *            chartitle7c: "",  //--
     *            chartitle7Image: {       //第七段图片
     *                chart7: [],
     *            },
     *            chartitle8: "",    // --
     *            chartitle8Image: {       //第八段图片
     *                chart8: [],
     *            },
     *            chartitle9: "",    //--
     *            chartitle9Image: {       //第八段图片
     *                chart9: [],
     *            },
     *       },
     *
     *
     * @param param
     * @return
     * @throws Exception
     */
    @RequestMapping(value = "downloadWord", method = RequestMethod.POST)
    @ResponseBody
    public ResponseEntity<byte[]> reportGeneration2(String param) throws Exception {

        ResponseEntity<byte[]> result = null;

            return this.buildResponseEntity(this.reportGeneration(param));

    }

    public static ResponseEntity<byte[]> buildResponseEntity(File file) throws IOException {
        if(file == null){
            return null;
        }
        byte[] body = null;
        String fileName = URLEncoder.encode(file.getName(),"utf-8");
        //获取文件
        InputStream is = new FileInputStream(file);
        body = new byte[is.available()];
        is.read(body);
        HttpHeaders headers = new HttpHeaders();
        //设置文件类型
        headers.add("Content-Disposition", "attchement;filename=" + fileName);
        //设置Http状态码
        HttpStatus statusCode = HttpStatus.OK;
        //返回数据
        ResponseEntity<byte[]> entity = new ResponseEntity<byte[]>(body, headers, statusCode);
        return entity;
    }


    public File reportGeneration(String reportDataStr) throws Exception {
        //对象转换
        JSONObject reportData = JSON.parseObject(reportDataStr);
        String title = (String) reportData.get("title");
        String time = (String) reportData.get("time1");
        //String author = (String) reportData.get("author");
        String fileName = time + title + ".docx";
        String wordPath = this.qcWordReportFolder+fileName;
        File wordFile = new File(wordPath);
        WordUtil wordUtil = WordUtil.builderWordUtil(this.qcWordReportFolder, fileName);
        //如果当前站点的报告已经生成过了即返回存在的文件
//        if(wordFile.exists()){
//            return wordFile;
//        }else{
        return this.generateReportWord3(wordUtil,wordFile,reportData);
        //  }
    }



    //拼接源解析报告
    private File generateReportWord3(WordUtil wordutil,File file,JSONObject reportData) throws Exception {
        //解析页面json
        //头部
        String title = (String) reportData.get("title");
        //时间
        String time = (String) reportData.get("titletwo");
        String time1 = (String) reportData.get("time1");
        String time2 = (String) reportData.get("time2");
        //第一段文字
        String firstTxt = (String) reportData.get("chartitle1");
        //表头
        String[] table1Title = {"  ","周一","周二","周三","周四","周五","周六","周日"};
        //获取第二段数据
        JSONObject secondMap = (JSONObject) reportData.get("chartitle1Form");
        //获取报告图表
        List<JSONArray> biaoge = (List<JSONArray>) secondMap.get("formt");

        //存放图片的路径
        String imageFolderStr = this.qcWordReportFolder+time+File.separator;

        File imageFolder = new File(imageFolderStr);

        //创建存放图片的文件夹
        if(!imageFolder.exists()){
            imageFolder.mkdirs();
        }
        //离子色谱文字
        String secondTxt = (String) reportData.get("chartitle2");
        String secondTxtc = (String) reportData.get("chartitle2c");
        //取第离子色谱图片
        JSONObject firstMap = (JSONObject) reportData.get("chartitle2Image");
        //图片
        String firstImageBase64Data = (String) firstMap.get("chart2");
        //将第一张图片的base64编码转换为图片存储到硬盘
        File firstImageFile = this.base64ToFileInputStream(new File(imageFolder+File.separator+System.currentTimeMillis()+"chartitle2Image.png"),firstImageBase64Data);

        //OCEC
        String secondTxt3 = (String) reportData.get("chartitle3");
        //取第离子色谱图片
        JSONObject firstMap3 = (JSONObject) reportData.get("chartitle3Image");
        //图片
        String firstImageBase64Data3 = (String) firstMap3.get("chart3");
        //将第一张图片的base64编码转换为图片存储到硬盘
        File firstImageFile3 = this.base64ToFileInputStream(new File(imageFolder+File.separator+System.currentTimeMillis()+"chartitle3Image.png"),firstImageBase64Data3);


        //重金属
        String secondTxt4 = (String) reportData.get("chartitle4");
        //取第离子色谱图片
        JSONObject firstMap4 = (JSONObject) reportData.get("chartitle4Image");
        //图片
        String firstImageBase64Data4 = (String) firstMap4.get("chart4");
        //将第一张图片的base64编码转换为图片存储到硬盘
        File firstImageFile4 = this.base64ToFileInputStream(new File(imageFolder+File.separator+System.currentTimeMillis()+"chartitle4Image.png"),firstImageBase64Data4);

        //颗粒物组分重构日变化
        String secondTxt5 = (String) reportData.get("chartitle5");
        //取第离子色谱图片
        JSONObject firstMap5 = (JSONObject) reportData.get("chartitle5Image");
        //图片
        String firstImageBase64Data5 = (String) firstMap5.get("chart5");
        //将第一张图片的base64编码转换为图片存储到硬盘
        File firstImageFile5 = this.base64ToFileInputStream(new File(imageFolder+File.separator+System.currentTimeMillis()+"chartitle5Image.png"),firstImageBase64Data5);

        //来源解析饼图
        String secondTxt6 = (String) reportData.get("chartitle6");
        //取第离子色谱图片
        JSONObject firstMap6 = (JSONObject) reportData.get("chartitle6Image");
        //图片
        String firstImageBase64Data6 = (String) firstMap6.get("chart6");
        //将第一张图片的base64编码转换为图片存储到硬盘
        File firstImageFile6 = this.base64ToFileInputStream(new File(imageFolder+File.separator+System.currentTimeMillis()+"chartitle6Image.png"),firstImageBase64Data6);

        //后向轨迹100
        String secondTxt7 = (String) reportData.get("chartitle7");
        String secondTxt7c = (String) reportData.get("chartitle7c");
        //取第离子色谱图片
        JSONObject firstMap7 = (JSONObject) reportData.get("chartitle7Image");
        //图片   下面  chartitle7Image  8Image  9Image  是已路径的形式保存到本地的，可以是网上任意路径都可以
        String firstImageBase64Data7 = (String) firstMap7.get("chart7");
        //File firstImageFile7 = this.base64ToFileInputStream(new File(imageFolder+File.separator+"chartitle7Image.png"),firstImageBase64Data7);
        this.savePage(firstImageBase64Data7,imageFolder+File.separator+"chartitle7Image.png");
        File firstImageFile7 = new File(imageFolder+File.separator+"chartitle7Image.png");
        //后向轨迹500
        String secondTxt8 = (String) reportData.get("chartitle8");
        //取第离子色谱图片
        JSONObject firstMap8 = (JSONObject) reportData.get("chartitle8Image");
        //图片
        String firstImageBase64Data8 = (String) firstMap8.get("chart8");
        //将第一张图片的base64编码转换为图片存储到硬盘
        //File firstImageFile8 = this.base64ToFileInputStream(new File(imageFolder+File.separator+"chartitle8Image.png"),firstImageBase64Data8);
        this.savePage(firstImageBase64Data8,imageFolder+File.separator+"chartitle8Image.png");
        File firstImageFile8 = new File(imageFolder+File.separator+"chartitle8Image.png");
        //后向轨迹1000
        String secondTxt9 = (String) reportData.get("chartitle9");
        //取第离子色谱图片
        JSONObject firstMap9 = (JSONObject) reportData.get("chartitle9Image");
        //图片
        String firstImageBase64Data9 = (String) firstMap9.get("chart9");
        //将第一张图片的base64编码转换为图片存储到硬盘
        //File firstImageFile9 = this.base64ToFileInputStream(new File(imageFolder+File.separator+"chartitle9Image.png"),firstImageBase64Data9);
        this.savePage(firstImageBase64Data9,imageFolder+File.separator+"chartitle9Image.png");
        File firstImageFile9 = new File(imageFolder+File.separator+"chartitle9Image.png");

        //瀑式函数 生成word
        boolean flag = wordutil
                //生成标题
                .createdBigTitle(title)
                .and()
                //word文字对部分   createdText可以设置文字样式大小   createdImage方法可以可以设置生成图片的大小
                .createdText(time, 12, false, ParagraphAlignment.CENTER,"仿宋_GB2312")
                .and()
                .createdText(time1, 8, false, ParagraphAlignment.CENTER,"仿宋_GB2312")
                .and()
                .createdText(time2, 10, false, ParagraphAlignment.CENTER,"仿宋_GB2312")
                .and()
                .createdText(firstTxt, 12, false, ParagraphAlignment.LEFT,"仿宋_GB2312")
                .and()
                //设置表格
                .setTableText(biaoge, table1Title)
                .and()
                //新的一页
                .setNewPage(BreakType.TEXT_WRAPPING)
                .and()
                //生成第一张图片
                .createdText(secondTxt, 12, false, ParagraphAlignment.LEFT,"仿宋_GB2312")
                .and()
                .createdText("        "+secondTxtc, 12, false, ParagraphAlignment.LEFT,"仿宋_GB2312")
                .and()
                .createdImage( "",firstImageFile)
                .and()
                //二
                .createdText(secondTxt3, 12, false, ParagraphAlignment.LEFT,"仿宋_GB2312")
                .and()
                .createdImage( "",firstImageFile3)
                //三
                .createdText(secondTxt4, 12, false, ParagraphAlignment.LEFT,"仿宋_GB2312")
                .and()
                .createdImage( "",firstImageFile4)
                //四
                .createdText(secondTxt5, 12, true, ParagraphAlignment.LEFT,"仿宋_GB2312")
                .and()
                .createdImage( "",firstImageFile5)
                //五
                .createdText(secondTxt6, 12, true, ParagraphAlignment.LEFT,"仿宋_GB2312")
                .and()
                .createdImage2( "",firstImageFile6)
                .and()
                //六
                .createdText(secondTxt7, 12, true, ParagraphAlignment.LEFT,"仿宋_GB2312")
                .and()
                .createdText("        "+secondTxt7c, 12, true, ParagraphAlignment.LEFT,"仿宋_GB2312")
                .and()
                .createdImage2( "",firstImageFile7)
                .and()
                //七
                .createdText(secondTxt8, 12, true, ParagraphAlignment.LEFT,"仿宋_GB2312")
                .and()
                .createdImage2( "",firstImageFile8)
                //八
                .createdText(secondTxt9, 12, true, ParagraphAlignment.LEFT,"仿宋_GB2312")
                .and()
                .createdImage2( "",firstImageFile9)
                .and()
                .setNewPage(BreakType.TEXT_WRAPPING)
                .and()
                //生成word 返回boolean值 true代表生成成功 false代表生成失败
                .closeDoc();
        if(flag){
            return file;
        }
        return null;
    }



    /**
     *TODO 解码base64  转化为图片
     * @param base64
     * @return
     */
    private File base64ToFileInputStream(File image,String base64) {
        FileOutputStream out = null;
        if(StringUtils.isEmpty(base64)){
            return null;
        }
        try {
            //判断图片是否已生成 存在则返回
            if(image.exists()){
                return image;
            }else{
                image.createNewFile();
                byte[] bytes = new BASE64Decoder().decodeBuffer(base64);
                //处理数据
                for(int i=0;i<bytes.length;i++){
                    if(bytes[i] < 0){
                        bytes[i] += 256;
                    }
                }

                out = new FileOutputStream(image);
                out.write(bytes);
                out.flush();
                return image;
            }

        } catch (IOException e) {
            e.printStackTrace();
        }finally {
            if(out!=null){
                try {
                    out.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return null;
    }


    /**
     * 保存图片到本地
     *
     */
    public  void savePage(String urlP,String path) throws Exception {
        //new一个URL对象
        URL url = new URL(urlP);
        //打开链接
        HttpURLConnection conn = (HttpURLConnection)url.openConnection();
        //设置请求方式为"GET"
        conn.setRequestMethod("GET");
        //超时响应时间为5秒
        conn.setConnectTimeout(5 * 1000);
        //通过输入流获取图片数据
        InputStream inStream = conn.getInputStream();
        //得到图片的二进制数据，以二进制封装得到数据，具有通用性
        byte[] data = readInputStream(inStream);
        //new一个文件对象用来保存图片，默认保存当前工程根目录
        File imageFile = new File(path);
        //创建输出流
        FileOutputStream outStream = new FileOutputStream(imageFile);
        //写入数据
        outStream.write(data);
        //关闭输出流
        outStream.close();
    }
    public static byte[] readInputStream(InputStream inStream) throws Exception{
        ByteArrayOutputStream outStream = new ByteArrayOutputStream();
        //创建一个Buffer字符串
        byte[] buffer = new byte[1024];
        //每次读取的字符串长度，如果为-1，代表全部读取完毕
        int len = 0;
        //使用一个输入流从buffer里把数据读取出来
        while( (len=inStream.read(buffer)) != -1 ){
            //用输出流往buffer里写入数据，中间参数代表从哪个位置开始读，len代表读取的长度
            outStream.write(buffer, 0, len);
        }
        //关闭输入流
        inStream.close();
        //把outStream里的数据写入内存
        return outStream.toByteArray();
    }





    /**
     * 下载 封装  excel不在本地生成  直接通过内存返回
     */
    public static void buildResponseEntity(HttpServletResponse response, ExcelUtils excelUtils) throws IOException {
        try {
            Workbook workBook = excelUtils.getWorkBook();
            String excelFileName = excelUtils.getExcelFileName();
            response.setContentType("multipart/form-data");
            response.setHeader("Content-Disposition", "attachment;filename="+ URLEncoder.encode(excelFileName, "utf-8"));
            OutputStream outputStream = response.getOutputStream();
            workBook.write(outputStream);
            outputStream.close();
            workBook.close();

        } catch (UnsupportedEncodingException e1) {
            e1.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
