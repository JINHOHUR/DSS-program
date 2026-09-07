package kr.co.dss.fx.web;

import jakarta.servlet.http.HttpServletRequest;
import kr.co.dss.fx.auth.UserContext;
import kr.co.dss.fx.service.SettingService;
import org.springframework.web.bind.annotation.ControllerAdvice;
import org.springframework.web.bind.annotation.ModelAttribute;

import java.time.LocalDate;
import java.time.format.TextStyle;
import java.util.Locale;

/** 모든 화면에 공통으로 넣는 값 */
@ControllerAdvice
public class GlobalModel {
    private final SettingService settings;
    public GlobalModel(SettingService settings) { this.settings = settings; }

    @ModelAttribute("me")
    public kr.co.dss.fx.auth.UserInfo me() { return UserContext.current(); }

    @ModelAttribute("companyName")
    public String companyName() { return settings.company(); }

    @ModelAttribute("today")
    public String today() {
        LocalDate d = LocalDate.now();
        return d + " (" + d.getDayOfWeek().getDisplayName(TextStyle.SHORT, Locale.KOREAN) + ")";
    }

    @ModelAttribute("path")
    public String path(HttpServletRequest req) { return req.getRequestURI(); }
}
