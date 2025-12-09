package com.example.poi_demo;

import com.deepoove.poi.policy.DynamicTableRenderPolicy;
import com.deepoove.poi.util.TableTools;
import org.apache.poi.xwpf.usermodel.*;

import java.util.List;
import java.util.Map;

/**
 * 扁平数据表格渲染插件（支持自动合并相同单元格）
 */
public class FlatTestItemTablePolicy extends DynamicTableRenderPolicy {

    @Override
    public void render(XWPFTable table, Object data) throws Exception {
        if (data == null) {
            return;
        }
        
        @SuppressWarnings("unchecked")
        List<Map<String, Object>> items = (List<Map<String, Object>>) data;
        
        if (items == null || items.isEmpty()) {
            return;
        }
        
        // 从第2行开始插入（第1行是表头）
        int startRow = 1;
        
        // 删除模板行（如果存在）
        if (table.getNumberOfRows() > 1) {
            table.removeRow(1);
        }
        
        // 插入所有数据行
        for (int i = 0; i < items.size(); i++) {
            XWPFTableRow row = table.insertNewTableRow(startRow + i);
            
            // 创建4列
            for (int j = 0; j < 4; j++) {
                row.createCell();
            }
            
            Map<String, Object> item = items.get(i);
            
            // 填充数据
            setCellText(row.getCell(0), getString(item, "seq"));
            setCellText(row.getCell(1), getString(item, "testItem"));
            setCellText(row.getCell(2), getString(item, "description"));
            setCellText(row.getCell(3), getString(item, "passed"));

            // 设置单元格样式
            setCellStyle(row.getCell(0), ParagraphAlignment.CENTER);
            setCellStyle(row.getCell(1), ParagraphAlignment.CENTER);
            setCellStyle(row.getCell(2), ParagraphAlignment.LEFT);
            setCellStyle(row.getCell(3), ParagraphAlignment.LEFT);
        }
        
        // 自动合并相同的"检测项"单元格（第2列，索引为1）
        mergeSameTestItems(table, items, startRow);
    }
    
    /**
     * 自动合并"检测项"列中相同值的单元格
     */
    private void mergeSameTestItems(XWPFTable table, List<Map<String, Object>> items, int startRow) {
        if (items.isEmpty()) {
            return;
        }
        
        int mergeStart = startRow;
        String currentTestItem = getString(items.get(0), "testItem");
        
        for (int i = 1; i < items.size(); i++) {
            String testItem = getString(items.get(i), "testItem");
            
            // 如果当前项与前一项不同，合并前面的单元格
            if (!testItem.equals(currentTestItem)) {
                int mergeEnd = startRow + i - 1;
                if (mergeEnd > mergeStart) {
                    // 合并第2列（索引为1）
                    TableTools.mergeCellsVertically(table, 1, mergeStart, mergeEnd);
                }
                // 更新合并起点和当前项
                mergeStart = startRow + i;
                currentTestItem = testItem;
            }
        }
        
        // 处理最后一组
        int mergeEnd = startRow + items.size() - 1;
        if (mergeEnd > mergeStart) {
            TableTools.mergeCellsVertically(table, 1, mergeStart, mergeEnd);
        }
    }
    
    /**
     * 设置单元格文本
     */
    private void setCellText(XWPFTableCell cell, String text) {
        if (text == null) {
            text = "";
        }
        
        //清空原有段落
        if (cell.getParagraphs().size() > 0) {
            for (int i = cell.getParagraphs().size() - 1; i >= 0; i--) {
                cell.removeParagraph(i);
            }
        }
        
        // 创建新段落
        XWPFParagraph paragraph = cell.addParagraph();
        paragraph.setAlignment(ParagraphAlignment.LEFT);
        
        // 创建运行并设置文本
        XWPFRun run = paragraph.createRun();
        run.setText(text);
        run.setFontSize(10);
        run.setFontFamily("宋体");
    }
    
    /**
     * 设置单元格样式（对齐方式、垂直居中）
     */
    private void setCellStyle(XWPFTableCell cell, ParagraphAlignment alignment) {
        if (cell.getParagraphs().size() > 0) {
            XWPFParagraph paragraph = cell.getParagraphs().get(0);
            paragraph.setAlignment(alignment);
        }
        
        // 设置单元格垂直居中
        cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
    }
    
    /**
     * 从 Map 中安全获取字符串
     */
    private String getString(Map<String, Object> map, String key) {
        Object value = map.get(key);
        return value != null ? value.toString() : "";
    }
}