package com.example.demo;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

@RequestMapping("test")
@RestController
public class TestController {

    @GetMapping("simple")
    public ModelAndView test() {
        Map<String, Object> map = new HashMap<>();
        List<User> list = new ArrayList<>();

        User user1 = new User();
        user1.setId("1");
        user1.setName("张三");
        user1.setAge(9);
        list.add(user1);

        User user2 = new User();
        user2.setId("2");
        user2.setName("李四");
        user2.setAge(12);
        list.add(user2);

        map.put("list", list);

        return new ModelAndView(new JxlsView("templates/test.xlsx", "output"), map);
    }

}
