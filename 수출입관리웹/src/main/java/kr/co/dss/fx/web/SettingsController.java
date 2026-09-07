package kr.co.dss.fx.web;

import kr.co.dss.fx.domain.Company;
import kr.co.dss.fx.service.MasterService;
import kr.co.dss.fx.service.SettingService;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import java.util.List;
import java.util.Map;

@Controller
@RequestMapping("/settings")
public class SettingsController {
    private final MasterService master;
    private final SettingService settings;
    public static final List<String> KINDS = List.of("고객사", "최종고객사", "공급사", "기타");
    public static final List<String> CODE_CATS = List.of("fx.구분", "com.통화", "fx.수량단위");

    public SettingsController(MasterService master, SettingService settings) { this.master = master; this.settings = settings; }

    @GetMapping
    public String page(Model m) {
        m.addAttribute("companies", master.allCompanies());
        m.addAttribute("kinds", KINDS);
        m.addAttribute("codeCats", CODE_CATS);
        java.util.Map<String, List<String>> codes = new java.util.LinkedHashMap<>();
        for (String c : CODE_CATS) codes.put(c, master.codes(c));
        m.addAttribute("codes", codes);
        m.addAttribute("lcRules", master.lcRules());
        m.addAttribute("nameMap", master.nameMap());
        m.addAttribute("values", settings.all());
        m.addAttribute("archiveRoot", settings.archiveRoot());
        m.addAttribute("active_menu", "settings");
        return "settings";
    }

    @PostMapping("/values")
    public String values(@RequestParam Map<String, String> p, RedirectAttributes ra) {
        for (String k : new String[]{"company", "ship_days", "expiry_days", "lc_lead", "comm_due", "horizon", "archive_path"})
            if (p.containsKey(k)) settings.set(k, p.get(k));
        ra.addFlashAttribute("msg", "설정을 저장했습니다");
        return "redirect:/settings";
    }

    @PostMapping("/company")
    public String company(@RequestParam String code, @RequestParam String name, @RequestParam(defaultValue = "") String shortName,
                          @RequestParam String kind, RedirectAttributes ra) {
        Company c = master.company(code.trim().toUpperCase()).orElseGet(Company::new);
        c.setCode(code.trim().toUpperCase()); c.setName(name.trim()); c.setShortName(shortName.trim()); c.setKind(kind);
        master.saveCompany(c);
        ra.addFlashAttribute("msg", "거래처 저장: " + c.getCode());
        return "redirect:/settings";
    }

    @PostMapping("/company/delete")
    public String companyDelete(@RequestParam String code, RedirectAttributes ra) {
        try { master.deleteCompany(code); ra.addFlashAttribute("msg", "거래처 삭제: " + code); }
        catch (Exception e) { ra.addFlashAttribute("err", "사용 중인 거래처는 삭제할 수 없습니다. 대신 '사용안함'으로 두세요."); }
        return "redirect:/settings";
    }

    @PostMapping("/code/add")
    public String codeAdd(@RequestParam String category, @RequestParam String value) {
        master.addCode(category, value); return "redirect:/settings";
    }

    @PostMapping("/code/delete")
    public String codeDelete(@RequestParam String category, @RequestParam String value) {
        master.deleteCode(category, value); return "redirect:/settings";
    }

    @PostMapping("/lcrule")
    public String lcRule(@RequestParam String companyCode, @RequestParam int shipDays, @RequestParam int expiryDays, RedirectAttributes ra) {
        master.saveLcRule(companyCode, shipDays, expiryDays);
        ra.addFlashAttribute("msg", "신용장 규칙 저장: " + companyCode + " +" + shipDays + "/+" + expiryDays + "일");
        return "redirect:/settings";
    }

    @PostMapping("/lcrule/delete")
    public String lcRuleDelete(@RequestParam String companyCode) { master.deleteLcRule(companyCode); return "redirect:/settings"; }
}
