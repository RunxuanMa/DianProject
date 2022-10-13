package com.Controller;

import com.service.ExcelService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import java.util.ArrayList;
import java.util.LinkedList;

@RestController
public class ViewControl {

    @Autowired
    private ExcelService excelService;

    @GetMapping(value = "/index")
    public ModelAndView index(){

        return new ModelAndView("/index");
    }

    @RequestMapping(value = "/myexcel")
    public ModelAndView myexcel(){
        return new ModelAndView("/myexcel");
    }

    @RequestMapping(value = "postexcel")
    public ModelAndView postexcel(){
        return new ModelAndView("/postexcel");
    }

    @RequestMapping(value = "getexcel")
    public ModelAndView getexcel(){
        ModelAndView mv = new ModelAndView();

        ArrayList<String> tableList;
        tableList=excelService.getTableMap();

        mv.addObject("table",tableList);
        mv.setViewName("/getexcel");
        return mv;
    }

}
