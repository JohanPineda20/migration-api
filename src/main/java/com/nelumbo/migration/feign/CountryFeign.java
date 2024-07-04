package com.nelumbo.migration.feign;

import com.nelumbo.migration.feign.dto.responses.CountryResponse;
import com.nelumbo.migration.feign.dto.responses.DefaultResponse;
import org.springframework.cloud.openfeign.FeignClient;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;

import java.util.List;

@FeignClient(name= "country", url="${hr-api}/countries")
public interface CountryFeign {
    @GetMapping
    DefaultResponse<List<CountryResponse>> findAll();

    @GetMapping("/{countryId}/states")
    DefaultResponse<List<CountryResponse>> findAllStatesByCountryId(@PathVariable Long countryId);

    @GetMapping("{countryId}/states/{stateId}/cities")
    DefaultResponse<List<CountryResponse>> findAllCitesByStateIdAndCountryId(@PathVariable Long countryId, @PathVariable Long stateId);
}