package com.moutsynov.app;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

public class Sportsman {

    private String firstname;
    private String lastname;
    private String patronymic;
    private Date birthday;

    public Sportsman(String fio, Date birthday) {
        String[] lfp = fio.split(" ");

        lastname = "";
        firstname = "";
        patronymic = "";

        if (lfp.length >= 1)
            lastname = fio.split(" ")[0];

        if (lfp.length >= 2)
            firstname = fio.split(" ")[1];

        if (lfp.length == 3)
            patronymic = fio.split(" ")[2];

        this.birthday = birthday;
    }

    // Возвращает строку для импорта в CSV файл
    public String getInfo() {
        DateFormat df = new SimpleDateFormat("dd.MM.yyyy");
        return String.format("%s;%s;%s;%s", lastname, firstname, patronymic, df.format(birthday));
    }

    // Генерирует INSERT для внесения информации в MSSQL
    public String getInsert() {
        DateFormat df = new SimpleDateFormat("yyyy-MM-dd");
        return String.format("INSERT INTO sportsmans (firstname, lastname, patronymic, birthday) VALUES('%s', '%s', '%s', '%s')",
                firstname.replace("'", "''"),
                lastname.replace("'", "''"),
                patronymic.replace("'", "''"),
                df.format(birthday));
    }

}
