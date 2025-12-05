package com.example.poi_demo;
 
import com.deepoove.poi.plugin.table.LoopRowTableRenderPolicy;
import com.deepoove.poi.util.TableTools;
import lombok.AllArgsConstructor;
import lombok.Data;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;
 
import java.util.Objects;
 
/**
 * @description: 单元格合并策略类
 */
@Data
@AllArgsConstructor
public class MergeTablePolicy extends LoopRowTableRenderPolicy {
    /**
     * 纵向合并列数（1开始）
     */
    private Integer verticallyMergeColumn;
    /**
     * 纵向合并开始行数(0开始)
     */
    private Integer verticallyMergeStartRow;
    /**
     * 横向合并列数（1开始）
     */
    private Integer horizontallyMergeColumn;
    /**
     * 横向合并开始行数(0开始)
     */
    private Integer horizontallyMergeStartRow;
 
    @Override
    protected void afterloop(XWPFTable table, Object data) {
        // 先处理纵向相邻相同的单元格合并
        if (verticallyMergeColumn != null && verticallyMergeStartRow != null) {
            mergeTableCellsVertically(table);
        }
        // 再处理横向相邻相同的单元格合并
        if (horizontallyMergeColumn != null && horizontallyMergeStartRow != null) {
            mergeTableCellsHorizontally(table);
        }
    }
 
 
    /**
     * @description:  合并纵向相同内容的单元格
     */
    private void mergeTableCellsVertically(XWPFTable table) {
        for (int col = 0; col < verticallyMergeColumn; col++) {
            // 前4行是表头，数据从第5行开始
            int mergeStart = verticallyMergeStartRow;
            String prevValue = null;
 
            for (int row = verticallyMergeStartRow; row < table.getNumberOfRows(); row++) {
                String currentValue = Objects.nonNull(table.getRow(row).getCell(col)) ? table.getRow(row).getCell(col).getText() : "";
 
                // 如果当前单元格不为空且与前一个值相等，则继续
                if (!currentValue.isBlank() && currentValue.equals(prevValue)) {
                    continue;
                }
 
                // 合并前一组单元格
                if (mergeStart < row - 1) {
                    TableTools.mergeCellsVertically(table, col, mergeStart, row - 1);
                }
 
                // 更新合并起点和前一个值
                mergeStart = row;
                prevValue = currentValue;
            }
 
            // 处理最后一组合并
            if (mergeStart < table.getNumberOfRows() - 1) {
                TableTools.mergeCellsVertically(table, col, mergeStart, table.getNumberOfRows() - 1);
            }
        }
    }
 
    /**
     * @description:  合并横向相同内容的单元
     */
    private void mergeTableCellsHorizontally(XWPFTable table) {
        for (int row = horizontallyMergeStartRow; row < table.getNumberOfRows(); row++) {
            int mergeStart = 0;
            String prevValue = null;
 
            for (int col = 0; col < horizontallyMergeColumn; col++) {
                String currentValue = Objects.nonNull(table.getRow(row).getCell(col)) ? table.getRow(row).getCell(col).getText() : "";
 
                // 如果当前单元格为空或与前一个值不同，则处理合并
                if (currentValue.isBlank() || !currentValue.equals(prevValue)) {
                    // 执行合并逻辑
                    if (mergeStart < col - 1) {
                        TableTools.mergeCellsHorizonal(table, row, mergeStart, col - 1);
                    }
                    // 更新起始合并位置和前一个值
                    mergeStart = col;
                    prevValue = currentValue;
                } else {
                    // 如果当前值等于前一个值，继续
                    continue;
                }
            }
 
            // 处理最后一组的合并
            if (mergeStart < horizontallyMergeColumn - 1) {
                TableTools.mergeCellsHorizonal(table, row, mergeStart, horizontallyMergeColumn - 1);
            }
        }
    }
}