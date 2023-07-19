package com.herambco.hrcapi.service.interfaces;

import org.springframework.web.multipart.MultipartFile;

import java.util.List;
import java.util.Set;

public interface FilterXLSInterface {
    List<Object> generateXLSBasedOnFilter(MultipartFile file, String itemFilter, Boolean gstFilter);

    Set<String> getItemList(MultipartFile file);
}
