package com.herambco.hrcapi.Dto;

import lombok.Data;
import org.springframework.web.multipart.MultipartFile;

import java.util.List;

@Data
public class XlsFilterRequestDto {
    String itemFilter;
    List<MultipartFile> files;
}
