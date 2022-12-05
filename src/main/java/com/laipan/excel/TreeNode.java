package com.laipan.excel;

import lombok.Data;

import java.util.List;

/**
 * <p><p/>
 *
 * @author laipan
 * @date 2022/12/02,9:50
 * @since v0.1
 */
@Data
public class TreeNode {

    private int depth;
    private int lft;
    private int rgt;
    private int rowNo;
    private int startCol;
    private int SameDepthNextNodeRowNo;

    private int childNodeNum;
    private int ord;
    private String expenseType;
    private List<NodeRecord> records;
    private List<TreeNode> children;
}
