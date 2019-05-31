package cn.quintin.crawl;

import org.apache.poi.hssf.record.aggregates.RowRecordsAggregate;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class MyCrawl {

    private static String outputFile = "d:test.xls";

    private static HSSFWorkbook workbook = null;


    public static void main(String[] args) throws IOException {


        String url = "https://s.weibo.com/top/summary?cate=realtimehot";
        Document doc = Jsoup.connect(url).userAgent("ie7：mozilla/4.0 (compatible; msie 7.0b; windows nt 6.0)")// 模拟浏览器访问
                .timeout(3000)// 设置超时
                .get();

        Element content = doc.getElementById("pl_feedlist_index");
        Elements links = content.getElementsByAttributeValue("action-type", "feed_list_item");

        //填入excel中进行展示
        workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("数据");
        int k = 0;
        HSSFRow row = sheet.createRow(k);
        row.createCell(0).setCellValue("微博id");
        row.createCell(1).setCellValue("微博昵称");
        row.createCell(2).setCellValue("微博内容");
        row.createCell(3).setCellValue("发布时间");
        row.createCell(4).setCellValue("发布平台");
        row.createCell(5).setCellValue("转发数");
        row.createCell(6).setCellValue("评论数");
        row.createCell(7).setCellValue("点赞数");
        for (Element link : links) {
            k++;
            row=sheet.createRow(k);
            try {
                //此处结合下面catch中的  continue; 实现对“热门文章”（和动态块里的标签种类不同，会执行catch中的代码）过滤；
                String id_href = link.getElementsByClass("name").first().attr("href");
                String pattern = "/(\\d{10})";
                Pattern r = Pattern.compile(pattern);
                Matcher m = r.matcher(id_href);
                String Id;

                if (m.find()) {
                    Id = m.group(1);
                } else {
                    Id = "";
                }


                String Name = link.getElementsByClass("name").text();
                String Content = link.getElementsByClass("txt").text();
                String Time = link.getElementsByClass("from").first().getElementsByAttributeValue("target", "_blank").text();
                String PlatForm = link.getElementsByClass("from").first().getElementsByAttributeValue("rel", "nofollow").text();
                String Forward = link.getElementsByClass("card-act").first().getElementsByTag("li").get(1).text();
                String Comment = link.getElementsByClass("card-act").first().getElementsByTag("li").get(2).text();
                String Like = link.getElementsByClass("card-act").first().getElementsByTag("li").get(3).text();


                row.createCell(0).setCellValue(Id);
                row.createCell(1).setCellValue(Name);
                row.createCell(2).setCellValue(Content);
                row.createCell(3).setCellValue(Time);
                row.createCell(4).setCellValue(PlatForm);
                row.createCell(5).setCellValue(Forward);
                row.createCell(6).setCellValue(Comment);
                row.createCell(7).setCellValue(Like);

            } catch (NullPointerException e) {
                continue;
            }
        }
        FileOutputStream fOut = new FileOutputStream(outputFile);
        workbook.write(fOut);
        fOut.flush();
        fOut.close();
        System.out.println("文件生成...");

    }

}
