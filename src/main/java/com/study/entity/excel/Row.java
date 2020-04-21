package com.study.entity.excel;

import java.util.List;

import lombok.Data;

@Data
public class Row {

    private Integer height;

    private List<Cell> cells;

    private Integer index;

}
