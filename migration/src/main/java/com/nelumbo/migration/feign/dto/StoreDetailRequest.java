package com.nelumbo.migration.feign.dto;

import lombok.Getter;
import lombok.Setter;

import java.util.ArrayList;
import java.util.List;
@Getter
@Setter
public class StoreDetailRequest {
    private List<Long> orgEntityDetailIds = new ArrayList<>();
}
