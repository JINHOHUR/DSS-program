package kr.co.dss.fx.web;

import kr.co.dss.fx.domain.Deal;
import kr.co.dss.fx.service.AlertService;
import kr.co.dss.fx.service.Dates;
import kr.co.dss.fx.service.DealService;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;

import java.math.BigDecimal;
import java.util.*;

@Controller
public class DashboardController {
    private final DealService deals;
    private final AlertService alerts;

    public DashboardController(DealService deals, AlertService alerts) { this.deals = deals; this.alerts = alerts; }

    public record CustRow(String name, int cnt, int late, int soon, String amount, String tag) {}

    @GetMapping("/")
    public String dashboard(Model m) {
        List<AlertService.Alert> al = alerts.compute();
        List<Deal> all = deals.all();
        List<Deal> active = all.stream().filter(d -> "진행중".equals(d.getStatus())).toList();
        long late = al.stream().filter(a -> "지연".equals(a.level())).count();
        long soon = al.stream().filter(a -> "임박".equals(a.level())).count();
        long comm = all.stream().filter(d -> d.getCommBilled() != null && d.getCommReceived() == null).count();

        Map<String, int[]> by = new LinkedHashMap<>();      // name -> [cnt, late, soon]
        Map<String, BigDecimal> amt = new HashMap<>();
        for (Deal d : active) {
            String k = d.customerName().isBlank() ? "(미지정)" : d.customerName();
            by.computeIfAbsent(k, x -> new int[3])[0]++;
            if (d.getAmount() != null) amt.merge(k, d.getAmount(), BigDecimal::add);
        }
        for (AlertService.Alert a : al) {
            String k = a.customer().isBlank() ? "(미지정)" : a.customer();
            int[] v = by.get(k);
            if (v == null) continue;
            if ("지연".equals(a.level())) v[1]++; else if ("임박".equals(a.level())) v[2]++;
        }
        List<CustRow> rows = new ArrayList<>();
        by.entrySet().stream().sorted((x, y) -> y.getValue()[0] - x.getValue()[0]).forEach(e -> {
            int[] v = e.getValue();
            rows.add(new CustRow(e.getKey(), v[0], v[1], v[2], Dates.money(amt.get(e.getKey())),
                    v[1] > 0 ? "late" : (v[2] > 0 ? "soon" : "")));
        });

        m.addAttribute("alerts", al);
        m.addAttribute("total", all.size());
        m.addAttribute("active", active.size());
        m.addAttribute("late", late);
        m.addAttribute("soon", soon);
        m.addAttribute("comm", comm);
        m.addAttribute("byCustomer", rows);
        m.addAttribute("active_menu", "dash");
        return "dashboard";
    }
}
