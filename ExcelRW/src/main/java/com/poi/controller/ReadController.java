package com.poi.controller;

import com.github.xiaoymin.knife4j.spring.annotations.EnableKnife4j;
import com.poi.service.ExcelService;
import com.result.Result;
import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;

/**
 * @author ak
 * @version 1.0
 * @date 2024/4/2 16:29
 */
@RestController
@RequestMapping("common/excel")
@Api("excel导入")
@Slf4j
@EnableKnife4j
public class ReadController {

    @Autowired
    private ExcelService excelService;

    @ApiOperation("excel导入")
    @PostMapping("/import")
    public Result importExcel(@RequestParam("file")MultipartFile file) throws IOException {
        return Result.success(excelService.importFile(file));
    }


}
