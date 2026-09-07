package kr.co.dss.fx.service;

import kr.co.dss.fx.Stages;
import kr.co.dss.fx.domain.Deal;
import kr.co.dss.fx.domain.StageLog;
import org.springframework.stereotype.Service;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.temporal.ChronoUnit;
import java.util.*;

/** 안건 목록 보기 5종 (전체 / 발주·납기 / 신용장 / 선적·서류 / 커미션) — spec/views.csv */
@Service
public class ViewService {
    private final DealService deals;
    private final MasterService master;
    private final SettingService settings;

    public ViewService(DealService deals, MasterService master, SettingService settings) {
        this.deals = deals; this.master = master; this.settings = settings;
    }

    public record Col(String key, String title, String align) {}
    public record Row(Long id, List<String> values, String tag) {}
    public record View(String key, String label, List<Col> cols) {}
    public record Result(View view, List<Row> rows, String summary) {}

    public static final List<View> VIEWS = List.of(
        new View("all", "전체", List.of(
            new Col("docNo", "관리번호", "c"), new Col("customer", "고객사", "l"), new Col("endUser", "최종고객사", "l"),
            new Col("model", "모델 / 품명", "l"), new Col("qty", "수량", "c"), new Col("po", "P.O 번호", "l"),
            new Col("stage", "진행단계", "l"), new Col("prog", "진척", "c"), new Col("fca", "FCA", "c"),
            new Col("next", "다음 기한", "l"), new Col("owner", "담당", "c"), new Col("status", "상태", "c"))),
        new View("order", "발주 · 납기 2~3", List.of(
            new Col("docNo", "관리번호", "c"), new Col("customer", "고객사", "l"), new Col("model", "모델 / 품명", "l"),
            new Col("qty", "수량", "c"), new Col("po", "P.O 번호", "l"), new Col("poDate", "P.O 접수", "c"),
            new Col("debit", "Debit Note", "l"), new Col("offer", "Offer Sheet", "l"), new Col("fca", "FCA", "c"),
            new Col("dday", "D-Day", "c"), new Col("s2", "2. P.O 접수", "c"), new Col("s3", "3. 공급사 주문", "c"),
            new Col("amount", "금액", "r"), new Col("cur", "통화", "c"))),
        new View("lc", "신용장 4", List.of(
            new Col("docNo", "관리번호", "c"), new Col("customer", "고객사", "l"), new Col("po", "P.O 번호", "l"),
            new Col("fca", "FCA", "c"), new Col("reqDue", "개설요청 기한", "c"), new Col("reqDday", "D-Day", "c"),
            new Col("bank", "개설은행", "l"), new Col("lcNo", "L/C 번호", "l"), new Col("open", "개설일", "c"),
            new Col("ship", "Latest shipment", "c"), new Col("exp", "Expiry date", "c"), new Col("expDday", "만료", "c"),
            new Col("rule", "적용규칙", "c"), new Col("amend", "Amend", "l"))),
        new View("ship", "선적 · 서류 5~7", List.of(
            new Col("docNo", "관리번호", "c"), new Col("customer", "고객사", "l"), new Col("po", "P.O 번호", "l"),
            new Col("fca", "FCA", "c"), new Col("ship", "Latest shipment", "c"), new Col("s5", "5. 서류 수령", "c"),
            new Col("s6", "6. 서류 발송", "c"), new Col("s7", "7. C/O 발급", "c"), new Col("exp", "L/C 만료", "c"),
            new Col("state", "상태", "l"), new Col("memo", "메모", "l"))),
        new View("comm", "커미션 8", List.of(
            new Col("docNo", "관리번호", "c"), new Col("customer", "고객사", "l"), new Col("po", "P.O 번호", "l"),
            new Col("amount", "납품금액", "r"), new Col("cur", "통화", "c"), new Col("rate", "요율%", "c"),
            new Col("camt", "커미션 금액", "r"), new Col("inv", "청구서 번호", "l"), new Col("billed", "청구일", "c"),
            new Col("passed", "경과", "c"), new Col("recv", "수령일", "c"), new Col("state", "상태", "c")))
    );

    public View view(String key) {
        return VIEWS.stream().filter(v -> v.key().equals(key)).findFirst().orElse(VIEWS.get(0));
    }

    public Result build(String viewKey, List<Deal> list) {
        View v = view(viewKey);
        List<Row> rows = new ArrayList<>();
        for (Deal d : list) {
            Set<Integer> done = deals.doneSet(d.getId());
            Map<Integer, StageLog> st = deals.stageMap(d.getId());
            rows.add(switch (v.key()) {
                case "order" -> rowOrder(d, done, st);
                case "lc" -> rowLc(d, done);
                case "ship" -> rowShip(d, done, st);
                case "comm" -> rowComm(d, done);
                default -> rowAll(d, done);
            });
        }
        return new Result(v, rows, summary(v.key(), rows));
    }

    // ---- 행 생성 ---------------------------------------------------------------
    private Row rowAll(Deal d, Set<Integer> done) {
        DealService.Progress p = deals.progress(d.getId());
        DealService.Deadline dl = deals.nextDeadline(d, done);
        String tag = d.isClosed() ? "done" : dl.tag();
        return new Row(d.getId(), List.of(d.getDocNo(), d.customerName(), d.endUserName(), nz(d.getModel()),
                d.qtyText(), nz(d.getPoNo()), p.nextText(), p.ratio(), Dates.iso(d.getFcaDate()), dl.text(),
                nz(d.getOwner()), d.getStatus()), tag);
    }

    private Row rowOrder(Deal d, Set<Integer> done, Map<Integer, StageLog> st) {
        Integer n = Dates.dday(d.getFcaDate());
        String tag = done.contains(5) ? "" : tagByDays(n);
        return new Row(d.getId(), List.of(d.getDocNo(), d.customerName(), nz(d.getModel()), d.qtyText(),
                nz(d.getPoNo()), Dates.iso(d.getPoDate()), nz(d.getDebitNo()), nz(d.getOfferNo()),
                Dates.iso(d.getFcaDate()), Dates.ddayText(n), doneDate(st, 2), doneDate(st, 3),
                Dates.money(d.getAmount()), nz(d.getCurrency())), tag);
    }

    private Row rowLc(Deal d, Set<Integer> done) {
        LocalDate due = deals.lcRequestDue(d);
        Integer n = Dates.dday(due);
        int[] r = master.lcRule(d.customerCode());
        boolean custom = master.hasCustomRule(d.customerCode());
        String tag = !done.contains(4) ? tagByDays(n) : (d.getLcNo() != null && !d.getLcNo().isBlank() ? "ok" : "");
        return new Row(d.getId(), List.of(d.getDocNo(), d.customerName(), nz(d.getPoNo()), Dates.iso(d.getFcaDate()),
                done.contains(4) ? "―" : Dates.iso(due), done.contains(4) ? "완료" : Dates.ddayText(n),
                nz(d.getLcBank()), nz(d.getLcNo()), Dates.iso(d.getLcOpenDate()), Dates.iso(d.getLatestShipment()),
                Dates.iso(d.getExpiryDate()), Dates.ddayText(Dates.dday(d.getExpiryDate())),
                "+" + r[0] + "/+" + r[1] + "일" + (custom ? " ★" : ""), nz(d.getLcAmend())), tag);
    }

    private Row rowShip(Deal d, Set<Integer> done, Map<Integer, StageLog> st) {
        String state, tag;
        if (done.contains(7)) { state = "서류 완료"; tag = "ok"; }
        else if (done.contains(6)) { state = "C/O 발급 대기"; tag = "soon"; }
        else if (done.contains(5)) { state = "고객사 발송 대기"; tag = "soon"; }
        else {
            Integer n = Dates.dday(d.getLatestShipment());
            if (n != null && n < 0) { state = "선적기한 경과"; tag = "late"; }
            else if (n != null && n <= 7) { state = "선적서류 수령 대기"; tag = "soon"; }
            else { state = "선적 전"; tag = ""; }
        }
        String memo = String.join(" / ", Arrays.asList(nz(st.get(5).getMemo()), nz(st.get(6).getMemo()), nz(st.get(7).getMemo()))
                .stream().filter(s -> !s.isBlank()).toList());
        return new Row(d.getId(), List.of(d.getDocNo(), d.customerName(), nz(d.getPoNo()), Dates.iso(d.getFcaDate()),
                Dates.iso(d.getLatestShipment()), doneDate(st, 5), doneDate(st, 6), doneDate(st, 7),
                Dates.iso(d.getExpiryDate()), state, memo), tag);
    }

    private Row rowComm(Deal d, Set<Integer> done) {
        int due = settings.commDue();
        String passed = "", state, tag;
        if (d.getCommReceived() != null) { state = "수령 완료"; tag = "ok"; }
        else if (d.getCommBilled() != null) {
            long n = ChronoUnit.DAYS.between(d.getCommBilled(), LocalDate.now());
            passed = n + "일";
            if (n >= due) { state = "입금 지연"; tag = "late"; } else { state = "입금 대기"; tag = "soon"; }
        } else if (done.contains(7)) { state = "청구 필요"; tag = "soon"; }
        else { state = "―"; tag = ""; }
        return new Row(d.getId(), List.of(d.getDocNo(), d.customerName(), nz(d.getPoNo()), Dates.money(d.getAmount()),
                nz(d.getCurrency()), d.getCommRate() == null ? "" : d.getCommRate().stripTrailingZeros().toPlainString(),
                Dates.money(d.getCommAmount()), nz(d.getCommInvoiceNo()), Dates.iso(d.getCommBilled()), passed,
                Dates.iso(d.getCommReceived()), state), tag);
    }

    // ---- 합계 ------------------------------------------------------------------
    private String summary(String key, List<Row> rows) {
        switch (key) {
            case "order": return rows.size() + "건    금액 합계 " + Dates.money(sum(rows, 12));
            case "lc": {
                long opened = rows.stream().filter(r -> !r.values().get(7).isBlank()).count();
                return rows.size() + "건    개설완료 " + opened + "건 / 미개설 " + (rows.size() - opened) + "건";
            }
            case "ship": return rows.size() + "건    서류 완료 " + rows.stream().filter(r -> "ok".equals(r.tag())).count() + "건";
            case "comm": {
                BigDecimal billed = BigDecimal.ZERO, recv = BigDecimal.ZERO;
                for (Row r : rows) {
                    BigDecimal amt = Dates.parseMoney(r.values().get(6));
                    if (amt == null) continue;
                    if (!r.values().get(8).isBlank()) billed = billed.add(amt);
                    if (!r.values().get(10).isBlank()) recv = recv.add(amt);
                }
                return rows.size() + "건    청구 " + Dates.money(billed) + "    수령 " + Dates.money(recv)
                        + "    미수 " + Dates.money(billed.subtract(recv));
            }
            default: return rows.size() + "건";
        }
    }

    private static BigDecimal sum(List<Row> rows, int idx) {
        BigDecimal t = BigDecimal.ZERO;
        for (Row r : rows) { BigDecimal v = Dates.parseMoney(r.values().get(idx)); if (v != null) t = t.add(v); }
        return t;
    }
    private static String doneDate(Map<Integer, StageLog> st, int no) {
        StageLog s = st.get(no);
        return s == null || s.getDoneDate() == null ? "―" : s.getDoneDate().toString();
    }
    private static String tagByDays(Integer n) {
        if (n == null) return "";
        return n < 0 ? "late" : (n <= 7 ? "soon" : "");
    }
    private static String nz(String s) { return s == null ? "" : s; }
}
