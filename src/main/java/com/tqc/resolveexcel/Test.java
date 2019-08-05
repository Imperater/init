package com.tqc.resolveexcel;


import java.text.ParseException;
import java.text.SimpleDateFormat;

public class Test {

    public static void main(String[] args) throws ParseException {
        SimpleDateFormat sdf = new SimpleDateFormat("hh:mm:ss");//如2016-08-10 20:40
        String simpleDateFormat = new SimpleDateFormat("10:00:00").toString();
        System.out.println(simpleDateFormat);
    }
}
