package kr.co.dss.fx.service;

import kr.co.dss.fx.domain.Setting;
import kr.co.dss.fx.repo.Repos.SettingRepo;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.util.LinkedHashMap;
import java.util.Map;

@Service
public class SettingService {
    private final SettingRepo repo;
    @Value("${dss.archive.root:}") private String archiveRootYml;

    public SettingService(SettingRepo repo) { this.repo = repo; }

    public String get(String key, String def) {
        return repo.findById(key).map(Setting::getValue).filter(v -> v != null && !v.isBlank()).orElse(def);
    }

    public int getInt(String key, int def) {
        try { return Integer.parseInt(get(key, String.valueOf(def)).trim()); }
        catch (NumberFormatException e) { return def; }
    }

    @Transactional
    public void set(String key, String value) { repo.save(new Setting(key, value == null ? "" : value.trim())); }

    public Map<String, String> all() {
        Map<String, String> m = new LinkedHashMap<>();
        for (String k : new String[]{"company", "ship_days", "expiry_days", "lc_lead", "comm_due", "horizon", "archive_path"})
            m.put(k, get(k, ""));
        return m;
    }

    /** 자료방 루트: DB 설정 > application.yml */
    public String archiveRoot() {
        String v = get("archive_path", "");
        return v.isBlank() ? archiveRootYml : v;
    }

    public int shipDays()   { return getInt("ship_days", 14); }
    public int expiryDays() { return getInt("expiry_days", 28); }
    public int lcLead()     { return getInt("lc_lead", 7); }
    public int commDue()    { return getInt("comm_due", 30); }
    public int horizon()    { return getInt("horizon", 21); }
    public String company() { return get("company", "㈜디에스에스"); }
}
