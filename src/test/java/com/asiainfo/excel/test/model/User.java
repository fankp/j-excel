package com.asiainfo.excel.test.model;

import com.asiainfo.excel.inter.ExcelDesc;

public class User {

    @ExcelDesc(cellName = "用户名", order = 2)
    private String userName;
    @ExcelDesc(cellName = "密码", order = 3)
    private String userPwd;
    @ExcelDesc(cellName = "用户ID", order = 1)
    private String userId;

    public String getUserName() {
        return userName;
    }

    public void setUserName(String userName) {
        this.userName = userName;
    }

    public String getUserPwd() {
        return userPwd;
    }

    public void setUserPwd(String userPwd) {
        this.userPwd = userPwd;
    }

    public String getUserId() {
        return userId;
    }

    public void setUserId(String userId) {
        this.userId = userId;
    }
}
