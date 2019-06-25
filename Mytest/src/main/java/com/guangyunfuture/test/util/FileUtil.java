package com.guangyunfuture.test.util;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.File;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Map;

public class FileUtil<T> {

    // 挂在项目的某文件夹下
    public static final String REPORT_PATH = "E:" + File.separator;


    /**
     * 导出数据
     *
     * @param fileName 文件名
     * @param titles   字段名
     * @param result   导出数据
     * @auther xuguocai
     */

    public static String exportExcel(String fileName, String[] titles, List<Map<String, Object>> result) {
        HSSFWorkbook wb;
        FileOutputStream fos;
        String tempName = fileName;
        try {
            Date date = new Date();
            SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd");
            fileName += "_" + df.format(date) + ".xls";

            fos = new FileOutputStream(FileUtil.REPORT_PATH + fileName);
            wb = new HSSFWorkbook();

            HSSFSheet sh = wb.createSheet();

            // 设置列宽
            for (int i = 0; i < titles.length - 1; i++) {
                sh.setColumnWidth(i, 256 * 15 + 184);
            }

            // 第一行表头标题，CellRangeAddress 参数：行 ,行, 列,列
            HSSFRow row = sh.createRow(0);
            HSSFCell cell = row.createCell(0);
            cell.setCellValue(new HSSFRichTextString(tempName));
//            cell.setCellStyle(fontStyle(wb));
            sh.addMergedRegion(new CellRangeAddress(0, 0, 0, titles.length - 1));

            // 第二行
            HSSFRow row3 = sh.createRow(1);

            // 第二行的列
            for (int i = 0; i < titles.length; i++) {
                cell = row3.createCell(i);
                cell.setCellValue(new HSSFRichTextString(titles[i]));
//                cell.setCellStyle(fontStyle(wb));
            }

            //填充数据的内容  i表示行，z表示数据库某表的数据大小，这里使用它作为遍历条件
            int i = 2, z = 0;
            while (z < result.size()) {
                row = sh.createRow(i);
                Map<String, Object> map = result.get(z);
                for (int j = 0; j < titles.length; j++) {
                    cell = row.createCell(j);
                    if (map.get(titles[j]) != null) {
                        cell.setCellValue(map.get(titles[j]).toString());
                    } else {
                        cell.setCellValue("");
                    }
                }
                i++;
                z++;
            }

            wb.write(fos);
            fos.flush();
            fos.close();
            return FileUtil.REPORT_PATH + fileName;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

}