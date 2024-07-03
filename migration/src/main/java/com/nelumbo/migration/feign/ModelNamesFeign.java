package com.nelumbo.migration.feign;

import com.nelumbo.migration.feign.dto.responses.DefaultResponse;
import com.nelumbo.migration.feign.dto.responses.ModelFieldsResponse;
import org.springframework.cloud.openfeign.FeignClient;
import org.springframework.web.bind.annotation.*;

import java.util.List;

@FeignClient(name= "modelNames", url="localhost:8080/core-api/v1/model-names")
public interface ModelNamesFeign {

    @GetMapping("/{keyword}}/model-fields")
    DefaultResponse<List<ModelFieldsResponse>> findModelFieldsByKeyword(
            @RequestHeader("Authorization") String token,
            @PathVariable String keyword);
}
