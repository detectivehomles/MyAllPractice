package com.guangyunfuture.test.controller;

import com.guangyunfuture.test.util.FileUtil;
import com.guangyunfuture.test.vo.ReturnVo;
import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;
import io.swagger.annotations.ApiResponse;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@RestController
@RequestMapping("poi")
@Slf4j
public class PoiController {

    @GetMapping("getExcel")
    @ApiOperation(value = "生成excel文件",httpMethod = "GET")
    public void getExcel(HttpServletResponse response) throws Exception {
        log.info("[start] poi--getExcel");
        ServletOutputStream outputStream = null;
        try {
            Workbook wb = new HSSFWorkbook(); // 定义一个新的工作簿
            Sheet sheet = wb.createSheet("第一个Sheet页");  // 创建第一个Sheet页
            Row row = sheet.createRow(0); // 创建一个行
            row.setHeightInPoints(30);//设置这一行的高度

            createCell(wb, row, (short) 0, HorizontalAlignment.CENTER, VerticalAlignment.BOTTOM);//既在中间又在下边
            createCell(wb, row, (short) 1, HorizontalAlignment.FILL, VerticalAlignment.CENTER);//要充满屏幕又要中间
            createCell(wb, row, (short) 2, HorizontalAlignment.LEFT, VerticalAlignment.TOP);//既在右边又在上边
            createCell(wb, row, (short) 3, HorizontalAlignment.RIGHT, VerticalAlignment.TOP);//既在右边又在上边
            for (int i = 1; i <= 100; i++) {
                Row row1 = sheet.createRow(i);
                for (int j = 0; j < 4; j++) {
                    Cell cell = row1.createCell(j);
                    cell.setCellValue(i);
                }
            }

//        FileOutputStream fileOut = new FileOutputStream("e:\\工作簿.xls");
//        wb.write(fileOut);
//        fileOut.close();
            String filename = "tmp.xls";
            response.setHeader("content-disposition", "attachment;filename="+filename);
            outputStream = response.getOutputStream();
            wb.write(outputStream);
        } catch (IOException e) {
            e.printStackTrace();
            log.error("[error] poi--getExcel");
        } finally {
            outputStream.close();
            log.info("[end] poi--getExcel");
        }
//        return new ReturnVo("已保存到 e:\\工作簿.xls");
    }

    /**
     * 创建一个单元格并为其设定指定的对其方式
     *  @param wb     工作簿
     * @param row    行
     * @param column 列
     * @param halign 水平方向对其方式
     * @param valign 垂直方向对其方式
     */
    private static void createCell(Workbook wb, Row row, short column, HorizontalAlignment halign, VerticalAlignment valign) {
        Cell cell = row.createCell(column);  // 创建单元格
        cell.setCellValue(new HSSFRichTextString("标题"));  // 设置值

        CellStyle cellStyle = wb.createCellStyle(); // 创建单元格样式

        cellStyle.setAlignment(halign);  // 设置单元格水平方向对其方式
        cellStyle.setVerticalAlignment(valign); // 设置单元格垂直方向对其方式

        cell.setCellStyle(cellStyle); // 设置单元格样式
    }


    @ApiOperation(value = "导出地址数据")
    @PostMapping(value="/exportAddresses")
    public ReturnVo exportAddresses(){
        // 数据库表对应的字段
        String[] titles = new String[] {"id","name","pid","code","description"};
        List<Map<String,Object>> objList = new ArrayList<>();

        // 数据库表对应的数据
//        List<AddressVo> list = addressService.exportAddresses();
//        for(AddressVo item : list){
//            Map<String,Object> tempMap = new HashMap<>();
//            tempMap.put("id", item.getId());
//            tempMap.put("name", item.getName());
//            tempMap.put("pid", item.getPid());
//            tempMap.put("code", item.getCode());
//            tempMap.put("description", item.getDescription());
//            objList.add(tempMap);
//        }

        for (int i = 0; i < 10; i++) {
            Map<String,Object> tempMap = new HashMap<>();
            tempMap.put("id", i);
            tempMap.put("name", "name"+i);
            tempMap.put("pid", "pid"+i);
            tempMap.put("code", "code"+i);
            tempMap.put("description", "description"+i);
            objList.add(tempMap);
        }
        String path = FileUtil.exportExcel("地址树",titles,objList);
        System.out.println("path="+path);
        ReturnVo resp = new ReturnVo(path);
        if(path == null){
            resp.setMsg("导出失败！");
        }

        return resp;
    }

}
