package com.Controller;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.ExceptionHandler;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.ModelAndView;

@RequestMapping
@Controller
public class ExceptionController {
    /**
     * 该方法返回ModelAndView：目的是为了可以让我们封装视图和错误信息
     * @param e 参数 Exception e:会将产生异常对象注入到方法中
     * @return
     */
    @ExceptionHandler(value = {java.lang.RuntimeException.class})
    public ModelAndView arithmeticExceptionHandler(Exception e){
        ModelAndView mv = new ModelAndView();
        mv.addObject("errorMsg",e);
        mv.setViewName("/error");
        return mv;
    }
}
