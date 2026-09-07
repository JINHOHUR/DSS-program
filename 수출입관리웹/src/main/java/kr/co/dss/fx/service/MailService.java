package kr.co.dss.fx.service;

import kr.co.dss.fx.auth.UserContext;
import kr.co.dss.fx.domain.Deal;
import org.springframework.stereotype.Service;

import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;

/** Offer Sheet 송부 메일 문안 (엑셀 'OfferSheet 송부' 시트 양식) */
@Service
public class MailService {
    private final MasterService master;
    private final SettingService settings;
    private final DealService deals;

    public MailService(MasterService master, SettingService settings, DealService deals) {
        this.master = master; this.settings = settings; this.deals = deals;
    }

    public String offerSheetMail(List<Deal> list) {
        if (list.isEmpty()) return "";
        Deal first = list.get(0);
        LocalDate fca = first.getFcaDate();
        String fcaTxt = fca == null ? "____" : fca.getMonthValue() + "월 " + fca.getDayOfMonth() + "일";
        String sender = UserContext.current().display();

        List<String> L = new ArrayList<>();
        L.add("수신 : " + or(first.customerName(), "고객사명") + " / " + or(first.getCustomerPic(), "담당자명"));
        L.add("발신 : " + settings.company() + " / " + or(sender, "담당자"));
        L.add("");
        L.add("1. 귀사의 일익 번창하심을 기원합니다.");
        L.add("");
        L.add("2. 선적 전 신용장 개설 관련입니다. (FCA " + fcaTxt + ")");
        L.add("   첨부한 Offer sheet 참고하시고 하기의 내용을 반영하시어");
        L.add("   FCA일정 일주일 전까지 신용장 개설 부탁드립니다.");
        L.add("");

        String[] hdr = {"NO", "PO", "Offer sheet", "FCA", "Latest shipment", "Expiry date"};
        List<String[]> rows = new ArrayList<>();
        int i = 1;
        for (Deal d : list) {
            int[] r = master.lcRule(d.customerCode());
            LocalDate ls = d.getLatestShipment() != null ? d.getLatestShipment() : (d.getFcaDate() == null ? null : d.getFcaDate().plusDays(r[0]));
            LocalDate ex = d.getExpiryDate() != null ? d.getExpiryDate() : (d.getFcaDate() == null ? null : d.getFcaDate().plusDays(r[1]));
            rows.add(new String[]{String.valueOf(i++), nz(d.getPoNo()), nz(d.getOfferNo()), Dates.iso(d.getFcaDate()),
                    ls == null ? "" : Dates.kdate(ls) + "( 선적일 기준 +" + (r[0] / 7) + "주)",
                    ex == null ? "" : Dates.kdate(ex) + " ( 선적일 기준 +" + (r[1] / 7) + "주)"});
        }
        int[] w = new int[hdr.length];
        for (int c = 0; c < hdr.length; c++) {
            w[c] = width(hdr[c]);
            for (String[] row : rows) w[c] = Math.max(w[c], width(row[c]));
        }
        L.add(fmt(hdr, w));
        StringBuilder sep = new StringBuilder("  ");
        for (int c = 0; c < w.length; c++) { if (c > 0) sep.append("-+-"); sep.append("-".repeat(w[c])); }
        L.add(sep.toString());
        for (String[] row : rows) L.add(fmt(row, w));
        L.add("");
        L.add("   ※ 개설은행 : " + or(first.getLcBank(), "(기재)"));
        LocalDate due = deals.lcRequestDue(first);
        L.add("   ※ 신용장 개설요청서는 " + (due == null ? "____" : due.toString()) + " 까지 회신 부탁드립니다.");
        L.add("");
        L.add("감사합니다.");
        return String.join("\n", L);
    }

    private static String fmt(String[] cells, int[] w) {
        StringBuilder sb = new StringBuilder("  ");
        for (int c = 0; c < cells.length; c++) {
            if (c > 0) sb.append(" | ");
            sb.append(cells[c]).append(" ".repeat(Math.max(0, w[c] - width(cells[c]))));
        }
        return sb.toString();
    }
    /** 고정폭 기준 표시 너비 — 한글·전각 2칸 */
    private static int width(String s) {
        int n = 0;
        for (char ch : s.toCharArray()) n += (ch >= 0x1100 && (ch <= 0x11FF || (ch >= 0x2E80 && ch <= 0xA4CF) || (ch >= 0xAC00 && ch <= 0xD7A3) || (ch >= 0xF900 && ch <= 0xFAFF) || (ch >= 0xFE30 && ch <= 0xFE4F) || (ch >= 0xFF00 && ch <= 0xFF60) || (ch >= 0xFFE0 && ch <= 0xFFE6))) ? 2 : 1;
        return n;
    }
    private static String nz(String s) { return s == null ? "" : s; }
    private static String or(String s, String d) { return s == null || s.isBlank() ? d : s; }
}
