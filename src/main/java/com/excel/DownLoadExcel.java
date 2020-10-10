package com.excel;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.util.*;

/**
 * @author chenlong
 * @version 1.0
 * @description
 * @date
 */
@Controller
public class DownLoadExcel {

    @RequestMapping("/down")
    public void downExcel(HttpServletResponse response) throws IOException {
        String[] titles=new String[]{"index","count"};
        List<Map<String,Object>> dataList=new ArrayList<>();
        for(int i=0;i<65588;i++){
            Map<String,Object> map=new HashMap<>();
            map.put("index",i);
            map.put("count",i*2);
            dataList.add(map);
        }
        Map<String,String> titlesDataMap=new HashMap<>();
        titlesDataMap.put("index","index");
        titlesDataMap.put("count","count");
        HSSFWorkbook wb = this.setData(titles, dataList, titlesDataMap, "测试");
        response.reset();
        response.setHeader("Content-disposition", "attachment; filename="+new String("测试".getBytes("utf-8"), "ISO8859-1")+".xls");
        response.setContentType("application/msexcel");
        OutputStream os=response.getOutputStream();
        wb.write(os);
        os.close();
    }

    private HSSFWorkbook setData(String[] titles, List<Map<String,Object>> dataList, Map<String,String> titlesDataMap,String name) {
        //创建工作簿
        HSSFWorkbook wb=new HSSFWorkbook();
        /**需要创建的sheet个数*/
        int sheetCount=1;
        //判断数据是否超出65535行,超出需要创建多个sheet
        if(dataList.size()>65535){
           sheetCount= dataList.size()/65535;
           if(dataList.size()%65535>0){
               sheetCount+=1;
           }
        }
        for (int i = 1; i <= sheetCount; i++) {
            HSSFSheet sheet=wb.createSheet(name+"("+i+")");
            //设置标题行
            HSSFRow titleRow=sheet.createRow(0);
            //标题行单元格
            for(int j=0;j<titles.length;j++) {
                HSSFCell createCell = titleRow.createCell(j);
                createCell.setCellValue(titles[j]);
            }
            int count=1;
            //重新设置数据集合
            int startIndex=(i-1)*65535;
            int endIndex=(i-1)*65535+65534;
            if(i==sheetCount){
                endIndex=dataList.size()-1;
            }
            List<Map<String,Object>> data=dataList.subList(startIndex,endIndex);
            //封装数据行
            for(Map<String,Object> item:data) {
                HSSFRow row=sheet.createRow(count);
                for(int j=0;j<titles.length;j++){
                    HSSFCell cell = row.createCell(j);
                    if(item.get(titlesDataMap.get(titles[j])) instanceof Double) {
                        cell.setCellValue(((double)item.get(titlesDataMap.get(titles[j]))));
                    }else if(item.get(titlesDataMap.get(titles[j])) instanceof Integer) {
                        cell.setCellValue((int)item.get(titlesDataMap.get(titles[j])));
                    }else if(item.get(titlesDataMap.get(titles[j])) instanceof String) {
                        cell.setCellValue((String)item.get(titlesDataMap.get(titles[j])));
                    }
                }
                count++;
            }
        }
	    return wb;
    }
}
