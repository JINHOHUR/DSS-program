package kr.co.dss.fx.web;

import kr.co.dss.fx.Stages;
import kr.co.dss.fx.domain.Deal;
import kr.co.dss.fx.service.*;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import java.time.LocalDate;
import java.util.*;

@Controller
@RequestMapping("/deals")
public class DealController {
    private final DealService deals;
    private final ViewService views;
    private final MasterService master;
    private final SettingService settings;
    private final MailService mail;
    private final ExcelService excel;
    private final ArchiveService archive;

    public DealController(DealService deals, ViewService views, MasterService master, SettingService settings,
                          MailService mail, ExcelService excel, ArchiveService archive) {
        this.deals = deals; this.views = views; this.master = master; this.settings = settings;
        this.mail = mail; this.excel = excel; this.archive = archive;
    }

    // ---- 목록 ----------------------------------------------------------------
    @GetMapping
    public String list(@RequestParam(defaultValue = "all") String view, @RequestParam(defaultValue = "") String q,
                       @RequestParam(defaultValue = "") String customer, @RequestParam(defaultValue = "진행중") String status,
                       Model m) {
        ViewService.Result r = views.build(view, deals.filtered(q, customer, status));
        m.addAttribute("result", r);
        m.addAttribute("views", ViewService.VIEWS);
        m.addAttribute("view", view); m.addAttribute("q", q); m.addAttribute("customer", customer); m.addAttribute("status", status);
        m.addAttribute("customers", master.companies("고객사"));
        m.addAttribute("statuses", Stages.STATUSES);
        m.addAttribute("active_menu", "deals");
        return "deals/list";
    }

    @GetMapping("/new")
    public String create() { return "redirect:/deals/" + deals.create().getId() + "?step=0"; }

    // ---- 상세 / 단계별 진행 -------------------------------------------------------
    @GetMapping("/{id}")
    public String detail(@PathVariable Long id, @RequestParam(required = false) Integer step,
                         @RequestParam(defaultValue = "flow") String tab, Model m) {
        Deal d = deals.get(id);
        int cur = step != null ? Math.max(0, Math.min(Stages.COUNT - 1, step)) : deals.currentStage(id);
        m.addAttribute("deal", d);
        m.addAttribute("form", DealForm.of(d));
        m.addAttribute("stages", Stages.ALL);
        m.addAttribute("stageMap", deals.stageMap(id));
        m.addAttribute("done", deals.doneSet(id));
        m.addAttribute("progress", deals.progress(id));
        m.addAttribute("deadline", deals.nextDeadline(d));
        m.addAttribute("cur", cur);
        m.addAttribute("tab", tab);
        m.addAttribute("lcRule", master.lcRule(d.customerCode()));
        m.addAttribute("lcCustom", master.hasCustomRule(d.customerCode()));
        m.addAttribute("lcRequestDue", deals.lcRequestDue(d));
        m.addAttribute("customers", master.companies("고객사"));
        m.addAttribute("endUsers", master.companies("최종고객사"));
        m.addAttribute("kinds", master.codes("fx.구분"));
        m.addAttribute("currencies", master.codes("com.통화"));
        m.addAttribute("units", master.codes("fx.수량단위"));
        m.addAttribute("statuses", Stages.STATUSES);
        m.addAttribute("stageFolder", archive.stageFolder(d, cur));
        m.addAttribute("active_menu", "deals");
        return "deals/form";
    }

    @PostMapping("/{id}")
    public String save(@PathVariable Long id, @ModelAttribute("form") DealForm form,
                       @RequestParam(defaultValue = "0") int step, @RequestParam(defaultValue = "flow") String tab,
                       RedirectAttributes ra) {
        Deal d = deals.get(id);
        form.applyTo(d, master);
        if (form.isRecalcLc()) deals.autoLc(d, true);
        deals.save(d);
        ra.addFlashAttribute("msg", "저장했습니다 — " + d.getDocNo());
        return "redirect:/deals/" + id + "?step=" + step + "&tab=" + tab;
    }

    /** 단계 저장: done=on 이면 완료(날짜 없으면 오늘). action=next 면 다음 단계로 이동 */
    @PostMapping("/{id}/stage/{no}")
    public String stage(@PathVariable Long id, @PathVariable int no,
                        @RequestParam(required = false) String done,
                        @RequestParam(required = false) String doneDate,
                        @RequestParam(defaultValue = "") String memo,
                        @RequestParam(defaultValue = "save") String action,
                        @ModelAttribute("form") DealForm form, RedirectAttributes ra) {
        Deal d = deals.get(id);
        form.applyTo(d, master);                    // 같은 화면의 입력값도 함께 저장
        if (form.isRecalcLc()) deals.autoLc(d, true);
        deals.save(d);

        LocalDate dd = (doneDate == null || doneDate.isBlank()) ? null : LocalDate.parse(doneDate);
        int next = no;
        if ("next".equals(action)) {
            next = deals.completeStage(id, no, memo);
            if (no == Stages.COUNT - 1) ra.addFlashAttribute("msg", "마지막 단계까지 완료했습니다. 상태를 '완료'로 바꾸려면 [전체 항목]에서 변경하세요.");
            else ra.addFlashAttribute("msg", no + "단계 완료 → " + next + ". " + Stages.name(next));
        } else {
            boolean isDone = done != null;
            deals.setStage(id, no, isDone ? (dd != null ? dd : LocalDate.now()) : null, memo);
            ra.addFlashAttribute("msg", "저장했습니다");
        }
        return "redirect:/deals/" + id + "?step=" + next;
    }

    @PostMapping("/{id}/delete")
    public String delete(@PathVariable Long id, RedirectAttributes ra) {
        deals.softDelete(id);
        ra.addFlashAttribute("msg", "삭제했습니다");
        return "redirect:/deals";
    }

    @PostMapping("/{id}/duplicate")
    public String duplicate(@PathVariable Long id) { return "redirect:/deals/" + deals.duplicate(id).getId(); }

    /** 단계별 자료 폴더로 이동 */
    @GetMapping("/{id}/folder")
    public String folder(@PathVariable Long id, @RequestParam(required = false) Integer stage) {
        String rel = archive.stageFolder(deals.get(id), stage);
        return "redirect:/archive?path=" + java.net.URLEncoder.encode(rel, java.nio.charset.StandardCharsets.UTF_8);
    }

    // ---- 신용장 자동계산 (JSON) --------------------------------------------------
    @GetMapping("/lc-calc")
    @ResponseBody
    public Map<String, Object> lcCalc(@RequestParam String fca, @RequestParam(defaultValue = "") String customer) {
        LocalDate f = LocalDate.parse(fca);
        int[] r = master.lcRule(customer);
        return Map.of("latestShipment", f.plusDays(r[0]).toString(), "expiryDate", f.plusDays(r[1]).toString(),
                "requestDue", f.minusDays(settings.lcLead()).toString(), "shipDays", r[0], "expiryDays", r[1],
                "custom", master.hasCustomRule(customer));
    }

    // ---- Offer Sheet 메일 ----------------------------------------------------
    @GetMapping("/mail")
    public String mail(@RequestParam(defaultValue = "") String ids, Model m) {
        List<Deal> list = new ArrayList<>();
        for (String s : ids.split(",")) if (!s.isBlank()) { try { list.add(deals.get(Long.parseLong(s.trim()))); } catch (Exception ignore) {} }
        m.addAttribute("text", mail.offerSheetMail(list));
        m.addAttribute("count", list.size());
        m.addAttribute("active_menu", "deals");
        return "deals/mail";
    }

    // ---- 엑셀 ----------------------------------------------------------------
    @GetMapping("/export.xlsx")
    public ResponseEntity<byte[]> export() throws Exception {
        byte[] body = excel.build();
        String name = "수출입안건_" + LocalDate.now().toString().replace("-", "") + ".xlsx";
        String enc = java.net.URLEncoder.encode(name, java.nio.charset.StandardCharsets.UTF_8).replace("+", "%20");
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename*=UTF-8''" + enc)
                .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                .body(body);
    }
}
