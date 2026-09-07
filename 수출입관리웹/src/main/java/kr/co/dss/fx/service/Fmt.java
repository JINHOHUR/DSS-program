package kr.co.dss.fx.service;

import org.springframework.stereotype.Component;

import java.math.BigDecimal;
import java.time.LocalDate;

/** 템플릿에서 ${@fmt.xxx(...)} 로 쓰는 표시 도우미 */
@Component("fmt")
public class Fmt {
    public String money(BigDecimal v) { return Dates.money(v); }
    public String date(LocalDate d) { return Dates.iso(d); }
    public String kdate(LocalDate d) { return Dates.kdate(d); }
    public String dday(LocalDate d) { return Dates.ddayText(Dates.dday(d)); }
    public String ddayN(Integer n) { return Dates.ddayText(n); }
    public String nz(String s) { return s == null ? "" : s; }
}
