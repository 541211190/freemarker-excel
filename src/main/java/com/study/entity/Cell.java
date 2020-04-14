package com.study.entity;

import com.study.dto.Data;

@lombok.Data
public class Cell {

    private String styleID;

    private Integer mergeAcross;

    private Integer MergeDown;

    private Data data;

    private Integer index;

}
