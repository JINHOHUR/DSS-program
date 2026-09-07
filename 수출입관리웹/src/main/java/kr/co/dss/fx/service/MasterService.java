package kr.co.dss.fx.service;

import kr.co.dss.fx.domain.CodeItem;
import kr.co.dss.fx.domain.Company;
import kr.co.dss.fx.domain.LcRule;
import kr.co.dss.fx.repo.Repos.CodeRepo;
import kr.co.dss.fx.repo.Repos.CompanyRepo;
import kr.co.dss.fx.repo.Repos.LcRuleRepo;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.util.*;
import java.util.stream.Collectors;

/** 거래처 · 선택목록 · 신용장 규칙 (기준정보) */
@Service
public class MasterService {
    private final CompanyRepo companies;
    private final CodeRepo codes;
    private final LcRuleRepo lcRules;
    private final SettingService settings;

    public MasterService(CompanyRepo companies, CodeRepo codes, LcRuleRepo lcRules, SettingService settings) {
        this.companies = companies; this.codes = codes; this.lcRules = lcRules; this.settings = settings;
    }

    // ---- 거래처 ------------------------------------------------------------
    public List<Company> companies(String kind) {
        return companies.findByKindOrderBySortAscNameAsc(kind).stream().filter(Company::isActive).collect(Collectors.toList());
    }
    public List<Company> allCompanies() { return companies.findAllByOrderByKindAscSortAscNameAsc(); }
    public Optional<Company> company(String code) { return code == null || code.isBlank() ? Optional.empty() : companies.findById(code); }
    public Map<String, String> nameMap() {
        Map<String, String> m = new HashMap<>();
        for (Company c : companies.findAll()) m.put(c.getCode(), c.getName());
        return m;
    }
    @Transactional public Company saveCompany(Company c) { return companies.save(c); }
    @Transactional public void deleteCompany(String code) { companies.deleteById(code); }

    // ---- 선택 목록 ---------------------------------------------------------
    public List<String> codes(String category) {
        return codes.findByCategoryOrderBySortAscValueAsc(category).stream().map(CodeItem::getValue).collect(Collectors.toList());
    }
    @Transactional
    public void addCode(String category, String value) {
        if (value == null || value.isBlank() || codes.findByCategoryAndValue(category, value).isPresent()) return;
        CodeItem c = new CodeItem();
        c.setCategory(category); c.setValue(value.trim());
        c.setSort(codes.findByCategoryOrderBySortAscValueAsc(category).size());
        codes.save(c);
    }
    @Transactional
    public void deleteCode(String category, String value) {
        codes.findByCategoryAndValue(category, value).ifPresent(codes::delete);
    }

    // ---- 신용장 규칙 --------------------------------------------------------
    public List<LcRule> lcRules() { return lcRules.findAll(); }

    /** (선적기한 가산일, 유효기간 가산일). 고객사 전용 규칙 없으면 기본값 */
    public int[] lcRule(String companyCode) {
        if (companyCode != null && !companyCode.isBlank()) {
            Optional<LcRule> r = lcRules.findById(companyCode);
            if (r.isPresent()) return new int[]{r.get().getShipDays(), r.get().getExpiryDays()};
        }
        return new int[]{settings.shipDays(), settings.expiryDays()};
    }
    public boolean hasCustomRule(String companyCode) {
        return companyCode != null && lcRules.existsById(companyCode);
    }
    @Transactional
    public void saveLcRule(String companyCode, int ship, int expiry) {
        LcRule r = new LcRule();
        r.setCompanyCode(companyCode); r.setShipDays(ship); r.setExpiryDays(expiry);
        lcRules.save(r);
    }
    @Transactional public void deleteLcRule(String companyCode) { lcRules.deleteById(companyCode); }
}
