package com.guangyunfuture.test.vo;

import lombok.Data;

@Data
public class ReturnVo {
    private String stauts;

    private String msg;

    private Object data;

    public ReturnVo (Object data){
        stauts = "200";
        msg = "success";
        this.data = data;
    }
}
