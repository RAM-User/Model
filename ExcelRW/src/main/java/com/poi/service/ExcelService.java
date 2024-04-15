package com.poi.service;

import org.springframework.web.multipart.MultipartFile;

/**
 * @author ak
 * @version 1.0
 * @date 2024/4/2 17:03
 */
public interface ExcelService {
    String importFile(MultipartFile file);
}
