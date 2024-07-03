package com.nelumbo.migration.feign.dto;

import lombok.Getter;
import lombok.Setter;

import java.util.List;
@Getter
@Setter
public class Page <T>{
    private List<T> content;
}
