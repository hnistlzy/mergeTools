package com.pactera;

import com.pactera.pojo.User;

import java.lang.reflect.Field;


public class TestMain {
    public static void main(String[] args)  {
        Class<User> userClass = User.class;
        String name = userClass.getName();
        System.out.println(name);

        Field[] fields = userClass.getDeclaredFields();


    }
}
