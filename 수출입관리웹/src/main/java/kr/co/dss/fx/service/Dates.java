package kr.co.dss.fx.service;

import java.math.BigDecimal;
import java.text.NumberFormat;
import java.time.LocalDate;
import java.time.temporal.ChronoUnit;
import java.util.Locale;

/** 날짜·금액 표시 유틸 (spec/rules.md §8) */
public final class Dates {
    private Dates() {}

    public static Integer dday(LocalDate d) {
        return d == null ? null : (int) ChronoUnit.DAYS.between(LocalDate.now(), d);
    }

    public static String ddayText(Integer n) {
        if (n == null) return "";
        if (n == 0) return "D-DAY";
        return n > 0 ? "D-" + n : "D+" + (-n);
    }

    public static String kdate(LocalDate d) {
        return d == null ? "" : d.getYear() + "년 " + d.getMonthValue() + "월 " + d.getDayOfMonth() + "일";
    }

    public static String iso(LocalDate d) { return d == null ? "" : d.toString(); }

    public static String money(BigDecimal v) {
        if (v == null) return "";
        NumberFormat f = NumberFormat.getNumberInstance(Locale.KOREA);
        f.setMaximumFractionDigits(v.stripTrailingZeros().scale() > 0 ? 2 : 0);
        return f.format(v);
    }

    /** "48,000,000" / "48000000" → BigDecimal. 숫자가 아니면 null */
    public static BigDecimal parseMoney(String s) {
        if (s == null) return null;
        String t = s.replace(",", "").trim();
        if (t.isEmpty()) return null;
        try { return new BigDecimal(t); } catch (NumberFormatException e) { return null; }
    }

    /** "지연" / "임박" / "예정" */
    public static String level(Integer n) {
        if (n == null) return "";
        return n < 0 ? "지연" : (n <= 7 ? "임박" : "예정");
    }
}
