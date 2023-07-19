package com.herambco.hrcapi.Dto;

import lombok.AllArgsConstructor;
import lombok.Data;

@Data
@AllArgsConstructor
public class XlsFilterResponseDto {
    String itemName;
    String qty;
    String gst;
    Double beforeGst;
    String amount;
}
