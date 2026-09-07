package kr.co.dss.fx.service;

import kr.co.dss.fx.domain.Deal;
import org.springframework.stereotype.Service;

import java.time.LocalDate;
import java.time.temporal.ChronoUnit;
import java.util.*;

/** 기한 알림 (spec/rules.md §6) */
@Service
public class AlertService {
    private final DealService dealService;
    private final SettingService settings;

    public AlertService(DealService dealService, SettingService settings) {
        this.dealService = dealService; this.settings = settings;
    }

    public record Alert(String level, Long dealId, String docNo, String customer, String po,
                        String item, LocalDate date, Integer dday, String detail) {
        public String ddayText() { return Dates.ddayText(dday); }
        public String dateText() { return Dates.iso(date); }
        public String tag() {
            return switch (level) { case "지연" -> "late"; case "임박" -> "soon"; case "확인" -> "ok"; default -> ""; };
        }
    }

    public List<Alert> compute() {
        int lead = settings.lcLead(), due = settings.commDue(), horizon = settings.horizon();
        List<Alert> out = new ArrayList<>();
        for (Deal d : dealService.all()) {
            if (d.isClosed()) continue;
            Set<Integer> done = dealService.doneSet(d.getId());

            if (!done.contains(4) && d.getFcaDate() != null) {
                LocalDate t = d.getFcaDate().minusDays(lead);
                Integer n = Dates.dday(t);
                if (n <= horizon) out.add(a(Dates.level(n), d, "L/C 개설요청 기한", t, n,
                        "FCA " + d.getFcaDate() + " 기준 " + lead + "일 전까지 개설요청서 수령 필요"));
            }
            if (!done.contains(5) && d.getFcaDate() != null) {
                Integer n = Dates.dday(d.getFcaDate());
                if (n <= horizon) out.add(a(Dates.level(n), d, "FCA (선적예정)", d.getFcaDate(), n, "선적서류 수령 준비 / 납기 Follow-up"));
            }
            if (!done.contains(5) && d.getLatestShipment() != null) {
                Integer n = Dates.dday(d.getLatestShipment());
                if (n <= horizon) out.add(a(Dates.level(n), d, "Latest shipment", d.getLatestShipment(), n, "신용장 선적기한 — 초과 시 Amend 필요"));
            }
            if (!done.contains(6) && d.getExpiryDate() != null) {
                Integer n = Dates.dday(d.getExpiryDate());
                if (n <= horizon) out.add(a(Dates.level(n), d, "L/C 유효기간 만료", d.getExpiryDate(), n, "선적서류 매입 / 발송 완료 확인"));
            }
            if (done.contains(7) && !done.contains(8) && d.getCommBilled() == null) {
                out.add(a("확인", d, "커미션 미청구", null, null, "원산지증명서까지 완료 — 커미션 청구서 송부 필요"));
            }
            if (d.getCommBilled() != null && d.getCommReceived() == null) {
                long passed = ChronoUnit.DAYS.between(d.getCommBilled(), LocalDate.now());
                if (passed >= due) {
                    LocalDate t = d.getCommBilled().plusDays(due);
                    out.add(a("지연", d, "커미션 미수령", t, Dates.dday(t),
                            d.getCommBilled() + " 청구 후 " + passed + "일 경과 — 매월 25일 전후 입금 확인"));
                }
            }
        }
        Map<String, Integer> rank = Map.of("지연", 0, "임박", 1, "확인", 2, "예정", 3);
        out.sort(Comparator.comparing((Alert x) -> rank.getOrDefault(x.level(), 9))
                .thenComparing(x -> x.dday() == null ? 9999 : x.dday()));
        return out;
    }

    private static Alert a(String level, Deal d, String item, LocalDate date, Integer n, String detail) {
        return new Alert(level, d.getId(), d.getDocNo(), d.customerName(), d.getPoNo() == null ? "" : d.getPoNo(),
                item, date, n, detail);
    }
}
